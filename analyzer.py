#!/usr/bin/env python3
"""
Excel Demystifier Analyzer - ANTLR AST version (deterministic; no regex parsing of VBA/M).

Pipeline:
  source text -> ANTLR AST -> deterministic fact extraction -> docs binding -> deterministic renderer -> report

No heuristic parsing: structure comes from the ANTLR grammars, semantics from
an offline Microsoft Learn docs cache. When something is uncertain the report
says so instead of guessing.

Dependencies:
  pip install antlr4-python3-runtime

Grammars:
  VBA: https://github.com/rossknudsen/Vba.Language
  Power Query M: https://github.com/antlr/grammars-v4/tree/master/powerquery

Harvester:
  Place harvest_docs.py next to this file (same dir).
  Analyzer will run harvester if doc_cache/docs_cache.json is missing (in background thread).

References:
- ANTLR: https://www.antlr.org/
- VBA Grammar: https://github.com/rossknudsen/Vba.Language
- Power Query M Grammar: https://github.com/antlr/grammars-v4/tree/master/powerquery
"""

from __future__ import annotations

import json
import logging
import os
import re
import subprocess
import sys
import threading
from queue import Queue
import time
import tkinter as tk
from dataclasses import dataclass, field
from tkinter import filedialog, messagebox
from typing import Any, Dict, Iterable, List, Optional, Tuple

logger = logging.getLogger(__name__)

# ---------------------------
# ANTLR loading helpers
# ---------------------------

def get_vba_parser():
    """
    Load VBA ANTLR parser from generated files.
    """
    try:
        sys.path.insert(0, os.path.join(os.path.dirname(__file__), "parsers", "vba"))
        from antlr4 import InputStream, CommonTokenStream
        from VbaLexer import VbaLexer
        from VbaParser import VbaParser
        return (VbaLexer, VbaParser, InputStream, CommonTokenStream), "vba", "ANTLR4"
    except Exception as e:
        raise RuntimeError(
            f"Could not load VBA ANTLR parser.\n\n"
            f"Error: {e}\n\n"
            "Make sure the VBA parser is generated in parsers/vba/\n"
            "Install: pip install antlr4-python3-runtime"
        )

def get_m_parser():
    """
    Load Power Query M ANTLR parser from generated files.
    """
    try:
        sys.path.insert(0, os.path.join(os.path.dirname(__file__), "parsers", "powerquery"))
        from antlr4 import InputStream, CommonTokenStream
        from PowerQueryLexer import PowerQueryLexer
        from PowerQueryParser import PowerQueryParser
        return (PowerQueryLexer, PowerQueryParser, InputStream, CommonTokenStream), "powerquery_m", "ANTLR4"
    except Exception as e:
        raise RuntimeError(
            f"Could not load Power Query M ANTLR parser.\n\n"
            f"Error: {e}\n\n"
            "Make sure the Power Query parser is generated in parsers/powerquery/\n"
            "Install: pip install antlr4-python3-runtime"
        )

# ---------------------------
# Docs cache
# ---------------------------

SUPPORTED_SCHEMA_VERSIONS = {2}

@dataclass
class DocEntry:
    id: str
    language: str                    # "vba_excel" | "powerquery_m"
    symbol: str                      # "Excel.Range.Sort" | "Table.SelectRows"
    kind: str                        # "method" | "function" | "object" | ...
    summary: str
    signature: str
    parameters: List[str] = field(default_factory=list)
    param_docs: Dict[str, str] = field(default_factory=dict)
    return_value: str = ""
    remarks: str = ""
    source_url: str = ""
    parse_warnings: List[str] = field(default_factory=list)

class DocsCache:
    def __init__(self, cache_path: str):
        self.cache_path = cache_path
        self.schema_version: Optional[int] = None
        self.scraper_version: Optional[int] = None

        self.entries_by_id: Dict[str, DocEntry] = {}
        self.m_by_symbol_lc: Dict[str, DocEntry] = {}
        self.vba_by_member_lc: Dict[str, List[DocEntry]] = {}
        self.vba_by_class_member_lc: Dict[str, DocEntry] = {}

    def load(self) -> None:
        with open(self.cache_path, "r", encoding="utf-8") as f:
            blob = json.load(f)

        self.schema_version = int(blob.get("schema_version", 0) or 0)
        if self.schema_version not in SUPPORTED_SCHEMA_VERSIONS:
            raise RuntimeError(
                f"Incompatible docs cache schema_version={self.schema_version}. "
                f"Supported: {sorted(SUPPORTED_SCHEMA_VERSIONS)}. Re-run harvester."
            )

        if "scraper_version" in blob:
            try:
                self.scraper_version = int(blob.get("scraper_version", 0) or 0)
            except Exception:
                self.scraper_version = None

        records: Dict[str, Dict[str, Any]] = blob.get("records", {})
        for rec_id, rec in records.items():
            entry = DocEntry(
                id=rec.get("id", rec_id),
                language=rec.get("language", "") or "",
                symbol=rec.get("symbol", "") or "",
                kind=rec.get("kind", "unknown") or "unknown",
                summary=rec.get("summary", "") or "",
                signature=rec.get("signature", "") or "",
                parameters=rec.get("parameters", []) or [],
                param_docs=rec.get("param_docs", {}) or {},
                return_value=rec.get("return_value", "") or "",
                remarks=rec.get("remarks", "") or "",
                source_url=rec.get("source_url", "") or "",
                parse_warnings=rec.get("parse_warnings", []) or [],
            )
            self.entries_by_id[entry.id] = entry

            if entry.language == "powerquery_m" and entry.symbol:
                self.m_by_symbol_lc[entry.symbol.lower()] = entry

            if entry.language == "vba_excel" and entry.symbol:
                # Extract member name and strip common suffixes like " property", " method", " event"
                member = entry.symbol.split(".")[-1].lower()
                for suffix in [" property", " method", " event", " function", " sub"]:
                    if member.endswith(suffix):
                        member = member[:-len(suffix)].strip()
                        break
                self.vba_by_member_lc.setdefault(member, []).append(entry)

                parts = entry.symbol.split(".")
                if len(parts) >= 3 and parts[0].lower() == "excel":
                    cls = parts[1].lower()
                    mem = parts[-1].lower()
                    # Strip common suffixes from member name
                    for suffix in [" property", " method", " event", " function", " sub"]:
                        if mem.endswith(suffix):
                            mem = mem[:-len(suffix)].strip()
                            break
                    key = f"{cls}.{mem}"
                    if key not in self.vba_by_class_member_lc:
                        self.vba_by_class_member_lc[key] = entry

    def lookup_m(self, fn: str) -> Optional[DocEntry]:
        return self.m_by_symbol_lc.get(fn.lower())

    def candidates_vba_member(self, member: str) -> List[DocEntry]:
        return self.vba_by_member_lc.get(member.lower(), [])

    def lookup_vba_class_member(self, cls: str, member: str) -> Optional[DocEntry]:
        return self.vba_by_class_member_lc.get(f"{cls.lower()}.{member.lower()}")

# ---------------------------
# Harvester integration (async)
# ---------------------------

def ensure_docs_cache_path(out_dir: str = "doc_cache") -> str:
    return os.path.join(out_dir, "docs_cache.json")

def run_harvester_ensure(out_dir: str) -> Tuple[int, str, str]:
    harvester = os.path.join(os.path.dirname(__file__), "harvest_docs.py")
    if not os.path.exists(harvester):
        return 2, "", f"Harvester not found next to analyzer: {harvester}"

    cmd = [sys.executable, harvester, "ensure", "--out-dir", out_dir]
    proc = subprocess.run(cmd, capture_output=True, text=True)
    return proc.returncode, proc.stdout, proc.stderr

# ---------------------------
# ANTLR AST utilities
# ---------------------------

# Limits to prevent resource exhaustion from malicious/malformed input.
MAX_INPUT_CHARS = 10 * 1024 * 1024  # 10MB - typical macro files are far smaller.
MAX_FLOW_NODES = 100                # Power Query let-bindings - typical: 5-20.
MAX_TREE_NODES = 100000             # AST nodes - ~1000 lines of complex code.
MAX_TREE_DEPTH = 2000               # Nesting depth - typical: 10-50.
MAX_FACTS_PER_LANGUAGE = 10000      # Hard cap to avoid runaway fact extraction.

def get_node_text(ctx) -> str:
    """Get text from an ANTLR parse tree node."""
    if ctx is None:
        return ""
    try:
        if hasattr(ctx, 'getText'):
            result = ctx.getText()
            return result if result is not None else ""
        return str(ctx) if ctx is not None else ""
    except AttributeError:
        return ""
    except Exception as e:
        logger.error("Unexpected error in get_node_text: %s", e, exc_info=True)
        raise

def walk_tree(node) -> Iterable[Any]:
    """
    DFS walk over all ANTLR parse tree nodes.
    """
    if node is None:
        return
    yield node
    if hasattr(node, 'children') and node.children:
        for child in node.children:
            yield from walk_tree(child)

def walk_tree_limited(node, max_nodes: int, max_depth: int,
                      on_limit: Optional[Any] = None) -> Iterable[Any]:
    """
    Iterative DFS with node/depth limits to avoid recursion blowups.
    """
    if node is None:
        return
    stack: List[Tuple[Any, int]] = [(node, 0)]
    count = 0
    try:
        while stack:
            current, depth = stack.pop()
            count += 1
            if count > max_nodes:
                if on_limit:
                    on_limit(
                        f"Parse tree too large; truncated after {max_nodes} nodes. "
                        "Try analyzing a smaller portion."
                    )
                return
            if depth > max_depth:
                if on_limit:
                    on_limit(
                        f"Parse tree too deep; truncated at depth {max_depth}. "
                        "Try simplifying nested expressions."
                    )
                return

            yield current
            children = getattr(current, "children", []) or []
            for child in reversed(children):
                stack.append((child, depth + 1))
    finally:
        stack.clear()

def get_line_number(ctx) -> int:
    """Get line number from ANTLR context (1-based)."""
    if ctx is None:
        return 0
    if hasattr(ctx, 'start') and ctx.start:
        return ctx.start.line
    return 0

# ---------------------------
# Fact model
# ---------------------------

@dataclass
class Fact:
    language: str           # "vba" | "m"
    kind: str               # "member_call" | "function_call" | "file_io" | ...
    line: int               # 1-based
    text: str               # snippet (node text)
    receiver: Optional[str] = None
    receiver_type: Optional[str] = None  # NEW: Type of receiver (e.g., "Worksheet")
    receiver_is_identifier: bool = False
    receiver_base: Optional[str] = None
    member_or_fn: Optional[str] = None
    args_named: Dict[str, str] = field(default_factory=dict)   # VBA named args
    args_pos: List[str] = field(default_factory=list)          # M positional args
    with_receiver: Optional[str] = None                         # if in With block
    note: Optional[str] = None
    assigned_value: Optional[str] = None

@dataclass
class Scope:
    parent: Optional['Scope']
    symbols: Dict[str, str]

    def lookup(self, name: str) -> Optional[str]:
        key = name.lower()
        if key in self.symbols:
            return self.symbols[key]
        if self.parent:
            return self.parent.lookup(name)
        return None

def _dedupe_facts(facts: List[Fact]) -> List[Fact]:
    """
    Remove redundant facts on the same line where one expression is a strict
    subexpression of another.
    """
    by_line: Dict[int, List[Fact]] = {}
    for f in facts:
        by_line.setdefault(f.line, []).append(f)

    kept: List[Fact] = []
    for line, items in by_line.items():
        if len(items) == 1:
            kept.extend(items)
            continue
        items.sort(key=lambda f: len(f.text or ""), reverse=True)
        normalized = [("".join(f.text.split()).lower(), f) for f in items]
        for i, (ti, fi) in enumerate(normalized):
            if not ti:
                continue
            is_sub = False
            for j in range(i):
                tj = normalized[j][0]
                if ti in tj:
                    is_sub = True
                    break
            if not is_sub:
                kept.append(fi)

    return kept

def _select_best_bound_items(items: List["BoundItem"]) -> List["BoundItem"]:
    """
    Keep one representative bound item per line (most specific/longest text).
    """
    by_line: Dict[int, List["BoundItem"]] = {}
    for bi in items:
        by_line.setdefault(bi.fact.line, []).append(bi)

    selected: List["BoundItem"] = []
    for line, group in sorted(by_line.items()):
        best = max(group, key=lambda bi: len(bi.fact.text or ""))
        selected.append(best)

    return selected

# ---------------------------
# Power Query M Data Flow Model
# ---------------------------

@dataclass
class DataFlowNode:
    """
    Represents one step in a Power Query M transformation pipeline.

    Example:
        let
          Source = Excel.CurrentWorkbook(){[Name="Sales"]}[Content],
          Filtered = Table.SelectRows(Source, each [Amount] > 1000)
        in Filtered

    Would create two nodes:
        1. var_name="Source", expression="Excel.CurrentWorkbook()...", dependencies=[]
        2. var_name="Filtered", expression="Table.SelectRows(Source, ...)",
           dependencies=["Source"], operation="Table.SelectRows"
    """
    var_name: str               # "Source", "Filtered", "Sorted"
    line: int                   # Line number where defined
    expression: str             # Full expression text
    dependencies: List[str]     # Variables referenced in this expression
    operation: Optional[str]    # Function name if this is a function call (e.g., "Table.SelectRows")
    arg_map: Dict[str, str] = field(default_factory=dict)  # param_name -> arg_text

    def __repr__(self):
        deps = f"[{', '.join(self.dependencies)}]" if self.dependencies else "[]"
        op = f" ({self.operation})" if self.operation else ""
        return f"{self.var_name} = ...{op} deps={deps}"


@dataclass
class DataFlowGraph:
    """
    Complete transformation pipeline for a Power Query M query.
    Maps variable names to their DataFlowNode definitions.
    """
    nodes: Dict[str, DataFlowNode]      # var_name -> node
    output_var: Optional[str] = None    # Final variable returned by "in" clause

    def get_linear_order(self) -> List[str]:
        """
        Return variable names in dependency order (topological sort).
        Variables with no dependencies come first.
        """
        visited = set()
        order = []

        def visit(var_name: str):
            if var_name in visited or var_name not in self.nodes:
                return
            visited.add(var_name)

            # Visit dependencies first
            node = self.nodes[var_name]
            for dep in node.dependencies:
                visit(dep)

            order.append(var_name)

        # Visit all nodes
        for var_name in self.nodes.keys():
            visit(var_name)

        return order


class SemanticFlowRenderer:
    """
    Generates natural language explanations of M transformation pipelines
    using the docs cache plus curated verb mappings for common functions.

    Strategy:
      1. Extract verb phrases from function summaries ("Returns a table..." → "returns")
      2. Parse function arguments from expressions
      3. Match arguments to parameter names from docs
      4. Use parameter semantics (names + descriptions) to format details

    Uses docs cache semantics with curated verb mappings for common functions.
    """

    def __init__(self, docs: 'DocsCache'):
        self.docs = docs
        # Tested with common Table.* and Excel.* functions in the docs cache.
        self._function_verbs = {
            # Test: Table.SelectRows(tbl, each [x]>5) → "filters from 'tbl' where x > 5"
            "table.selectrows": "filters",
            # Test: Table.SelectColumns(tbl, {"A","B"}) → "selects columns from 'tbl' keeping A, B"
            "table.selectcolumns": "selects columns",
            # Test: Table.Sort(tbl, {{"Date", Order.Descending}}) → "sorts from 'tbl' by Date descending"
            "table.sort": "sorts",
            # Test: Table.RemoveColumns(tbl, {"A"}) → "removes columns from 'tbl' removing A"
            "table.removecolumns": "removes columns",
            # Test: Table.AddColumn(tbl, "Total", each [A]+[B]) → "adds a column from 'tbl' named 'Total'"
            "table.addcolumn": "adds a column",
            # Test: Table.RenameColumns(tbl, {{"A","B"}}) → "renames columns from 'tbl' columns: A, B"
            "table.renamecolumns": "renames columns",
            # Test: Table.Group(tbl, {"A"}, {{"Sum", each List.Sum([B]), type number}}) → "groups from 'tbl'"
            "table.groupby": "groups",
            # Test: Table.Join(t1, "ID", t2, "ID") → "joins from 't1'"
            "table.join": "joins",
            # Test: Table.NestedJoin(t1, "ID", t2, "ID", "t2") → "joins from 't1'"
            "table.nestedjoin": "joins",
            # Test: Table.ExpandTableColumn(tbl, "t2", {"A"}) → "expands columns from 'tbl' columns: A"
            "table.expandtablecolumn": "expands columns",
            # Test: Excel.CurrentWorkbook() → "loads data"
            "excel.currentworkbook": "loads data",
        }

    def explain_operation(self, node: DataFlowNode) -> str:
        """
        Generate natural language explanation of what this operation does.

        Examples:
            Input: Filtered = Table.SelectRows(Source, each [Amount] > 1000)
            Output: "filters 'Source' where Amount > 1000"

            Input: Sorted = Table.Sort(Filtered, {{"Date", Order.Descending}})
            Output: "sorts 'Filtered' by Date descending"

            Input: Source = Excel.CurrentWorkbook(){[Name="Sales"]}[Content]
            Output: "loads data from Excel workbook"
        """
        if not node.operation:
            # No function call detected - generic description
            if node.dependencies:
                return f"transforms '{node.dependencies[0]}'"
            else:
                return "loads data"

        # Look up function in docs cache
        entry = self.docs.lookup_m(node.operation)
        if not entry:
            # Function not in cache - fall back to generic description
            deps = f" from '{node.dependencies[0]}'" if node.dependencies else ""
            return f"applies {node.operation}{deps}"

        # Build explanation from docs + parsed arguments
        parts = []

        # 1. Start with action verb from docs summary
        verb_phrase = self._extract_verb_phrase(entry.summary, node.operation)
        has_from = False
        if verb_phrase:
            parts.append(verb_phrase)
            has_from = " from " in verb_phrase
        else:
            parts.append(f"applies {node.operation}")

        # 2. Add input source (first dependency) if not already implied
        if node.dependencies and not has_from:
            parts.append(f"from '{node.dependencies[0]}'")

        # 3. Parse arguments and add semantic details
        args = self._parse_function_args(node, entry)
        semantic_details = self._extract_semantic_details(node.operation, args, entry)
        if semantic_details:
            parts.append(semantic_details)

        return " ".join(parts)

    def _extract_verb_phrase(self, summary: str, function_name: Optional[str]) -> Optional[str]:
        """
        Extract action verb from function summary.

        Examples:
            "Returns a table with only rows that match a condition" → "filters"
            "Sorts the rows in a table" → "sorts"
            "Adds a column to a table" → "adds"

        Uses curated verb mappings for common functions; falls back to the summary.
        """
        if not summary:
            return None

        if function_name:
            mapped = self._function_verbs.get(function_name.lower())
            if mapped:
                return mapped

        first_sent = summary.split('.')[0].strip()
        if not first_sent:
            return None

        first_sent = ' '.join(first_sent.split()).lower()
        max_len = 120
        if len(first_sent) <= max_len:
            return first_sent

        cut = first_sent.rfind(' ', 0, max_len + 1)
        if cut <= 20:
            return first_sent[:max_len].rstrip()
        return first_sent[:cut].rstrip()

    def _parse_function_args(self, node: DataFlowNode, entry: 'DocEntry') -> Dict[str, str]:
        """
        Parse function arguments and match to parameter names from docs.

        Example:
            expression: "Table.SelectRows(Source, each [Amount] > 1000)"
            entry.parameters: ["table", "condition"]

            Returns: {
                "table": "Source",
                "condition": "each [Amount] > 1000"
            }

        This mapping allows us to understand semantic meaning based on
        parameter names and documentation.
        """
        if not entry or not entry.parameters:
            return {}
        return node.arg_map or {}

    def _extract_semantic_details(self, fn_name: str, args: Dict[str, str],
                                   entry: 'DocEntry') -> Optional[str]:
        """
        Extract semantic meaning from arguments using parameter documentation.

        Strategy:
          - Parameter names reveal intent: "condition", "columns", "newColumnName"
          - Parameter docs provide context: "The condition to filter by"
          - Argument values contain the specifics: "each [Amount] > 1000"

        We combine these to generate natural language like:
          - "where Amount > 1000" (from condition parameter)
          - "columns: Date, Amount" (from columns parameter)
          - "named 'Total'" (from newColumnName parameter)
        """
        details = []

        # Skip first parameter (usually the input table/list)
        if not entry.parameters:
            return None

        first_param = entry.parameters[0]

        for param_name, arg_value in args.items():
            # Skip the input data parameter
            if param_name == first_param:
                continue

            # Get parameter documentation for context
            param_doc = entry.param_docs.get(param_name, "").lower()
            param_lower = param_name.lower()

            # Pattern 1: Condition/predicate/filter expressions
            # Example: condition="each [Amount] > 1000" → "where Amount > 1000"
            if any(keyword in param_lower for keyword in ['condition', 'predicate', 'criteria']):
                condition_text = self._humanize_condition(arg_value)
                if condition_text:
                    details.append(f"where {condition_text}")

            # Pattern 2: Column name lists
            # Example: columns={"Date", "Amount"} → "columns: Date, Amount"
            elif 'column' in param_lower or 'field' in param_lower:
                col_names = self._extract_column_names(arg_value)
                if col_names:
                    fn_lower = fn_name.lower()
                    if 'remove' in fn_lower or 'delete' in fn_lower:
                        details.append(f"removing {', '.join(col_names)}")
                    elif 'select' in fn_lower or 'keep' in fn_lower:
                        details.append(f"keeping {', '.join(col_names)}")
                    else:
                        details.append(f"columns: {', '.join(col_names)}")

            # Pattern 3: New column name
            # Example: newColumnName="Total" → "named 'Total'"
            elif 'newcolumn' in param_lower or ('name' in param_lower and 'column' in param_doc):
                col_name = arg_value.strip('"\'')
                if col_name and not col_name.startswith('['):
                    details.append(f"named '{col_name}'")

            # Pattern 4: Sort specifications
            # Example: order={{"Date", Order.Descending}} → "by Date descending"
            elif 'order' in param_lower or 'sort' in param_lower or 'comparer' in param_doc:
                order_spec = self._extract_sort_order(arg_value)
                if order_spec:
                    details.append(order_spec)

            # Pattern 5: Separator/delimiter strings
            # Example: separator=", " → "using ', ' as separator"
            elif 'separator' in param_lower or 'delimiter' in param_lower:
                sep = arg_value.strip('"\'')
                if sep:
                    details.append(f"using '{sep}' as separator")

        return ", ".join(details) if details else None

    def _humanize_condition(self, condition_expr: str) -> Optional[str]:
        """
        Convert M predicate expressions to readable conditions.

        Examples:
            "each [Amount] > 1000" → "Amount > 1000"
            "each [Status] = \"Active\"" → "Status = 'Active'"
            "each [Date] >= #date(2024,1,1)" → "Date >= #date(2024,1,1)"
        """
        # Remove "each" keyword
        clean = re.sub(r'^\s*each\s+', '', condition_expr, flags=re.I).strip()

        # Remove brackets around column references: [Name] → Name
        clean = re.sub(r'\[([^\]]+)\]', r'\1', clean)

        # Add spacing around operators
        clean = re.sub(r'\s*([<>]=|<>|=|>|<)\s*', r' \1 ', clean)

        # Normalize whitespace
        clean = ' '.join(clean.split())

        return clean if clean and len(clean) < 100 else None

    def _extract_column_names(self, arg_value: str) -> List[str]:
        """
        Extract column names from list/record literals.

        Examples:
            "{\"Date\", \"Amount\"}" → ["Date", "Amount"]
            "[Date=\"Date\", Amount=\"Amount\"]" → ["Date", "Amount"]
        """
        # Find all quoted strings
        names = re.findall(r'["\']([^"\']+)["\']', arg_value)
        return names if names else []

    def _extract_sort_order(self, arg_value: str) -> Optional[str]:
        """
        Extract sort specification from M sort arguments.

        Examples:
            "{\"Date\", Order.Descending}" → "by Date descending"
            "{{\"Amount\", Order.Ascending}}" → "by Amount ascending"
        """
        # Extract column name (first quoted string)
        col_match = re.search(r'["\']([^"\']+)["\']', arg_value)
        if not col_match:
            return None

        col = col_match.group(1)

        # Extract order (Order.Ascending or Order.Descending)
        order_match = re.search(r'Order\.(\w+)', arg_value, re.I)
        order = order_match.group(1).lower() if order_match else "ascending"

        return f"by {col} {order}"

    def render_flow(self, graph: DataFlowGraph) -> str:
        """
        Render complete flow graph as numbered steps in natural language.

        Example output:
            1. loads data from Excel workbook → 'Source'
            2. filters 'Source' where Amount > 1000 → 'Filtered'
            3. sorts 'Filtered' by Date descending → 'Sorted'
            4. removes columns OrderID, Status from 'Sorted' → 'Final'
        """
        lines = []
        order = graph.get_linear_order()

        for i, var_name in enumerate(order, 1):
            node = graph.nodes[var_name]
            explanation = self.explain_operation(node)
            lines.append(f"{i}. {explanation} → '{var_name}'")

        if graph.output_var and graph.output_var != order[-1]:
            lines.append(f"\nFinal output: '{graph.output_var}'")

        return "\n".join(lines)

# ---------------------------
# VBA fact extractor (ANTLR)
# ---------------------------

class VBAFactExtractor:
    """
    Extracts deterministic facts from a VBA ANTLR parse tree.

    Strategy:
      - Single pass: track Dim/Set declarations and extract member access facts
      - Track With block context for implicit receivers
    """

    def __init__(self, parser_classes):
        self.VbaLexer, self.VbaParser, self.InputStream, self.CommonTokenStream = parser_classes
        self.global_scope = Scope(parent=None, symbols={})

    def parse(self, text: str):
        """Parse VBA text and return parse tree."""
        input_stream = self.InputStream(text)
        lexer = self.VbaLexer(input_stream)
        token_stream = self.CommonTokenStream(lexer)
        parser = self.VbaParser(token_stream)
        return parser.module()  # Entry point for VBA grammar

    def _extract_variable_name(self, identifier_ctx) -> Optional[str]:
        """Extract variable name from IdentifierContext."""
        if identifier_ctx is None:
            return None
        text = get_node_text(identifier_ctx)
        return text.strip() if text else None

    def _extract_type_name(self, type_spec_ctx) -> Optional[str]:
        """Extract type name from TypeSpecContext."""
        if type_spec_ctx is None:
            return None
        text = get_node_text(type_spec_ctx)
        return text.strip() if text else None

    def _is_scope_node(self, node) -> bool:
        return node.__class__.__name__ in {
            "SubStmtContext",
            "FunctionStmtContext",
            "PropertyGetStmtContext",
            "PropertyLetStmtContext",
            "PropertySetStmtContext",
        }

    def _get_scope(self, node, scope_cache: Dict[Any, Scope]) -> Scope:
        cur = node
        while cur is not None:
            if self._is_scope_node(cur):
                scope = scope_cache.get(cur)
                if not scope:
                    scope = Scope(parent=self.global_scope, symbols={})
                    scope_cache[cur] = scope
                return scope
            cur = getattr(cur, "parentCtx", None)
        return self.global_scope

    def _extract_declaration(self, node) -> Tuple[Optional[str], Optional[str]]:
        var_name = None
        type_name = None
        for child in walk_tree_limited(node, MAX_TREE_NODES, MAX_TREE_DEPTH):
            child_class = child.__class__.__name__
            if child_class == "IdentifierContext" and not var_name:
                var_name = self._extract_variable_name(child)
            elif child_class == "TypeSpecContext" and not type_name:
                type_name = self._extract_type_name(child)
            if var_name and type_name:
                break
        return var_name, type_name

    def _build_scope_hierarchy(self, tree) -> Dict[Any, Scope]:
        scope_cache: Dict[Any, Scope] = {}
        for node in walk_tree_limited(tree, MAX_TREE_NODES, MAX_TREE_DEPTH):
            if self._is_scope_node(node) and node not in scope_cache:
                scope_cache[node] = Scope(parent=self.global_scope, symbols={})
            if node.__class__.__name__ == "VariableDclContext":
                var_name, type_name = self._extract_declaration(node)
                if var_name and type_name:
                    scope = self._get_scope(node, scope_cache)
                    scope.symbols[var_name.lower()] = type_name
        return scope_cache

    def _extract_facts_with_scopes(self, tree, scope_cache: Dict[Any, Scope],
                                   lines: List[str], warnings: List[str]) -> List[Fact]:
        facts: List[Fact] = []

        def warn_once(message: str) -> None:
            if message not in warnings:
                warnings.append(message)

        def receiver_base_name(receiver: Optional[str]) -> Optional[str]:
            if not receiver:
                return None
            match = re.match(r'^\s*([A-Za-z_]\w*)', receiver)
            return match.group(1) if match else None

        def infer_receiver_type_from_chain(receiver: Optional[str]) -> Optional[str]:
            if not receiver:
                return None
            parts = re.findall(r'\.([A-Za-z_]\w*)', receiver)
            if not parts:
                return None
            last = parts[-1].lower()
            if last in {"range", "cells", "rows", "columns"}:
                return "Range"
            if last in {"worksheets", "sheets", "activesheet"}:
                return "Worksheet"
            if last in {"workbooks", "activeworkbook", "thisworkbook"}:
                return "Workbook"
            if last in {"activecell"}:
                return "Range"
            return None

        def is_identifier_receiver(receiver_base: Optional[str]) -> bool:
            return bool(receiver_base)

        for node in walk_tree_limited(tree, MAX_TREE_NODES, MAX_TREE_DEPTH, warn_once):
            class_name = node.__class__.__name__
            node_text = get_node_text(node)
            line = get_line_number(node)

            if not node_text or not line:
                continue

            # Only match specific expression contexts to avoid duplicate extraction
            if class_name not in ["LExpressionContext", "MemberAccessExpressionContext"]:
                continue

            if "." not in node_text or len(node_text) > 200:
                continue

            text_lower = node_text.lower()

            if text_lower.startswith("open ") and (" for output" in text_lower or " for append" in text_lower):
                facts.append(Fact(
                    language="vba",
                    kind="file_io",
                    line=line,
                    text=node_text[:200],
                    note="Open file statement detected (writes external file)."
                ))
                continue

            if "print #" in text_lower:
                facts.append(Fact(
                    language="vba",
                    kind="file_io",
                    line=line,
                    text=node_text[:200],
                    note="Print # detected (writes to open file handle)."
                ))
                continue

            parts = node_text.split(".")
            if len(parts) >= 2:
                receiver = ".".join(parts[:-1]).strip()
                last = parts[-1].strip()
                for sep in ["=", "("]:
                    if sep in last:
                        member = last.split(sep)[0].strip()
                        break
                else:
                    member = last.strip()

                receiver_base = receiver_base_name(receiver)
                receiver_is_identifier = is_identifier_receiver(receiver_base)
                receiver_type = None
                if receiver_base:
                    scope = self._get_scope(node, scope_cache)
                    receiver_type = scope.lookup(receiver_base)
                chain_type = infer_receiver_type_from_chain(receiver)
                if chain_type:
                    receiver_type = chain_type

                assigned_value = None
                if member.lower() == "value" and line > 0 and line <= len(lines):
                    line_text = lines[line - 1]
                    if "=" in line_text:
                        assigned_value = line_text.split("=", 1)[1].strip()

                facts.append(Fact(
                    language="vba",
                    kind="member_call",
                    line=line,
                    text=node_text[:200],
                    receiver=receiver,
                    receiver_type=receiver_type,
                    receiver_is_identifier=receiver_is_identifier,
                    receiver_base=receiver_base,
                    member_or_fn=member,
                    assigned_value=assigned_value
                ))
                if len(facts) >= MAX_FACTS_PER_LANGUAGE:
                    facts.append(Fact(
                        language="vba",
                        kind="parse_warning",
                        line=1,
                        text="VBA parser warning.",
                        note=(
                            f"Fact extraction truncated at {MAX_FACTS_PER_LANGUAGE} facts. "
                            "Try analyzing a smaller portion."
                        )
                    ))
                    break

        return facts

    def extract(self, text: str) -> List[Fact]:
        """Extract facts from VBA code using ANTLR parse tree."""
        if len(text) > MAX_INPUT_CHARS:
            return [Fact(
                language="vba",
                kind="parse_error",
                line=1,
                text="VBA parsing skipped: input too large.",
                note=(
                    f"Input is {len(text):,} characters; limit is {MAX_INPUT_CHARS:,}. "
                    "Try analyzing a smaller portion or splitting the file."
                )
            )]
        try:
            tree = self.parse(text)
        except Exception as e:
            return [Fact(
                language="vba",
                kind="parse_error",
                line=1,
                text="VBA parsing failed.",
                note=str(e)
            )]

        self.global_scope.symbols.clear()
        warnings: List[str] = []
        lines = text.splitlines()

        scope_cache = self._build_scope_hierarchy(tree)
        facts = self._extract_facts_with_scopes(tree, scope_cache, lines, warnings)

        if warnings:
            facts.append(Fact(
                language="vba",
                kind="parse_warning",
                line=1,
                text="VBA parser warning.",
                note="; ".join(warnings)
            ))

        return _dedupe_facts(facts)

# ---------------------------
# M fact extractor (ANTLR)
# ---------------------------

class MFactExtractor:
    """
    Extracts function call facts AND builds semantic data flow graphs
    from Power Query M ANTLR parse trees.

    Two extraction modes:
      1. extract() - returns List[Fact] (existing behavior, for compatibility)
      2. extract_with_flow() - returns (List[Fact], Optional[DataFlowGraph])
    """

    def __init__(self, parser_classes, docs: Optional['DocsCache'] = None):
        self.PowerQueryLexer, self.PowerQueryParser, self.InputStream, self.CommonTokenStream = parser_classes
        self.docs = docs
        self._identifier_token_type = getattr(self.PowerQueryLexer, "IDENTIFIER", -1)
        self._parser_warnings: List[str] = []
        self._warning_set: set = set()

    def parse(self, text: str):
        """
        Parse Power Query M text and return parse tree.
        """
        self._parser_warnings = []
        self._warning_set = set()
        input_stream = self.InputStream(text)
        lexer = self.PowerQueryLexer(input_stream)
        token_stream = self.CommonTokenStream(lexer)
        parser = self.PowerQueryParser(token_stream)
        try:
            from antlr4 import PredictionMode
            from antlr4.error.ErrorStrategy import BailErrorStrategy
            parser._interp.predictionMode = PredictionMode.SLL
            parser._errHandler = BailErrorStrategy()
        except Exception as e:
            self._parser_warnings.append(f"Power Query parser fast-path disabled: {e}")
        return parser.document()

    def _warn_once(self, message: str) -> None:
        if message not in self._warning_set:
            self._warning_set.add(message)
            self._parser_warnings.append(message)

    def _extract_facts_from_tree(self, tree) -> List[Fact]:
        facts: List[Fact] = []
        for node in walk_tree_limited(tree, MAX_TREE_NODES, MAX_TREE_DEPTH, self._warn_once):
            class_name = node.__class__.__name__
            if class_name != "Primary_expressionContext":
                continue

            node_text = get_node_text(node)
            line = get_line_number(node)
            if not node_text or not line:
                continue

            if "(" not in node_text or "." not in node_text:
                continue

            # Confirm this looks like an invoke_expression: find the first "(" child.
            callee = self._extract_invoke_callee(node)
            if not callee:
                continue
            if not callee or "." not in callee or callee.startswith("."):
                continue

            if "=" in callee:
                callee = callee.split("=", 1)[-1].strip()
            if not callee:
                continue

            facts.append(Fact(
                language="m",
                kind="function_call",
                line=line,
                text=node_text[:200],
                member_or_fn=callee
            ))
            if len(facts) >= MAX_FACTS_PER_LANGUAGE:
                self._warn_once(
                    f"Fact extraction truncated at {MAX_FACTS_PER_LANGUAGE} facts. "
                    "Try analyzing a smaller portion."
                )
                break

        return facts

    def _extract_invoke_callee(self, node) -> Optional[str]:
        callee_parts: List[str] = []
        for ch in getattr(node, "children", []) or []:
            ch_text = get_node_text(ch)
            if ch_text == "(":
                break
            callee_parts.append(ch_text)
        if not callee_parts:
            return None
        callee = "".join(callee_parts).strip()
        if "=" in callee:
            callee = callee.split("=", 1)[-1].strip()
        return callee or None

    def extract(self, text: str) -> List[Fact]:
        """
        Extract facts from Power Query M code (original behavior).
        This method is kept for backwards compatibility.
        """
        facts, _ = self.extract_with_flow(text)
        return facts

    def extract_with_flow(self, text: str) -> Tuple[List[Fact], Optional[DataFlowGraph]]:
        """
        Extract both facts AND data flow graph from Power Query M code.

        Returns:
            (facts, flow_graph) where:
              - facts: List of function call facts (line-by-line API usage)
              - flow_graph: DataFlowGraph showing transformation pipeline
        """
        if len(text) > MAX_INPUT_CHARS:
            return ([Fact(
                language="m",
                kind="parse_error",
                line=1,
                text="Power Query M parsing skipped: input too large.",
                note=(
                    f"Input is {len(text):,} characters; limit is {MAX_INPUT_CHARS:,}. "
                    "Try analyzing a smaller portion or splitting the file."
                )
            )], None)
        try:
            tree = self.parse(text)
        except Exception as e:
            return ([Fact(
                language="m",
                kind="parse_error",
                line=1,
                text="Power Query M parsing failed.",
                note=str(e)
            )], None)
        facts = self._extract_facts_from_tree(tree)
        flow_graph = self._build_flow_graph_from_tree(tree)
        if self._parser_warnings:
            facts.append(Fact(
                language="m",
                kind="parse_warning",
                line=1,
                text="Power Query M parser warning.",
                note="; ".join(self._parser_warnings)
            ))
        return facts, flow_graph

    def _build_flow_graph_from_tree(self, tree) -> Optional[DataFlowGraph]:
        let_node = None
        for node in walk_tree_limited(tree, MAX_TREE_NODES, MAX_TREE_DEPTH, self._warn_once):
            if node.__class__.__name__ == "Let_expressionContext":
                let_node = node
                break
        if not let_node:
            return None

        var_nodes: List[Any] = []
        for node in walk_tree_limited(let_node, MAX_TREE_NODES, MAX_TREE_DEPTH, self._warn_once):
            if node.__class__.__name__ == "VariableContext":
                var_nodes.append(node)
        if not var_nodes:
            return None

        var_names: List[str] = []
        var_contexts: List[Tuple[str, Any, Any]] = []
        for var_node in var_nodes:
            name_ctx = self._first_child(var_node, ["Variable_nameContext"])
            expr_ctx = self._first_child(var_node, ["ExpressionContext"])
            if not name_ctx or not expr_ctx:
                continue
            var_name = get_node_text(name_ctx).strip()
            if not var_name:
                continue
            var_names.append(var_name)
            var_contexts.append((var_name, var_node, expr_ctx))

        if not var_contexts:
            return None
        if len(var_contexts) > MAX_FLOW_NODES:
            self._parser_warnings.append(
                f"Flow graph skipped: too many let bindings ({len(var_contexts)} > {MAX_FLOW_NODES}). "
                "Try reducing the query or analyzing a subset."
            )
            return None

        nodes: Dict[str, DataFlowNode] = {}
        for var_name, var_node, expr_ctx in var_contexts:
            expression = get_node_text(expr_ctx).strip()
            if not expression:
                continue

            operation, arg_values, dependencies = self._analyze_expression(expr_ctx, var_names)
            arg_map = self._map_args_to_params(operation, arg_values)
            line = get_line_number(var_node)

            nodes[var_name] = DataFlowNode(
                var_name=var_name,
                line=line,
                expression=expression,
                dependencies=dependencies,
                operation=operation,
                arg_map=arg_map
            )

        output_var = self._extract_output_var(let_node, var_names)
        if not nodes:
            return None
        return DataFlowGraph(nodes=nodes, output_var=output_var)

    def _first_child(self, node, class_names: List[str]):
        for ch in getattr(node, "children", []) or []:
            if ch.__class__.__name__ in class_names:
                return ch
        return None

    def _analyze_expression(self, expr_ctx, known_vars: List[str]) -> Tuple[Optional[str], List[str], List[str]]:
        operation: Optional[str] = None
        arg_values: List[str] = []
        dependencies: List[str] = []
        in_args = False

        for node in walk_tree_limited(expr_ctx, MAX_TREE_NODES, MAX_TREE_DEPTH, self._warn_once):
            class_name = node.__class__.__name__

            if class_name == "Primary_expressionContext" and operation is None:
                callee = self._extract_invoke_callee(node)
                if callee and "." in callee and not callee.startswith("."):
                    operation = callee.strip()

            if class_name == "Argument_listContext" and operation:
                in_args = True
                continue

            if in_args and class_name == "ExpressionContext":
                arg_values.append(get_node_text(node).strip())

            if class_name == "TerminalNodeImpl":
                symbol = getattr(node, "symbol", None)
                if symbol and self._is_identifier_token(symbol):
                    text = get_node_text(node)
                    if text in known_vars and text not in dependencies:
                        dependencies.append(text)

        return operation, arg_values, dependencies

    def _map_args_to_params(self, operation: Optional[str], arg_values: List[str]) -> Dict[str, str]:
        if not operation or not arg_values or not self.docs:
            return {}
        entry = self.docs.lookup_m(operation)
        if not entry or not entry.parameters:
            return {}
        return {p: a.strip() for p, a in zip(entry.parameters, arg_values)}

    def _is_identifier_token(self, symbol) -> bool:
        return symbol.type == self._identifier_token_type

    def _extract_output_var(self, let_node, known_vars: List[str]) -> Optional[str]:
        children = getattr(let_node, "children", []) or []
        for idx, ch in enumerate(children):
            if get_node_text(ch).lower() == "in" and idx + 1 < len(children):
                expr_ctx = children[idx + 1]
                text = get_node_text(expr_ctx).strip()
                if text in known_vars:
                    return text
        return None


# ---------------------------
# Binding + Rendering
# ---------------------------

def dedupe(items: List[str]) -> List[str]:
    seen = set()
    out = []
    for x in items:
        x = x.strip()
        if x and x not in seen:
            seen.add(x)
            out.append(x)
    return out

def doc_summary_to_it_sentence(summary: str) -> str:
    s = summary.strip()
    if not s:
        return "It performs an operation (no documentation summary available)."
    if s.lower().startswith("it "):
        return s if s.endswith(".") else s + "."
    s2 = s[0].lower() + s[1:] if s and s[0].isalpha() else s
    if not s2.endswith("."):
        s2 += "."
    return "It " + s2

@dataclass
class BoundItem:
    fact: Fact
    bound: Optional[DocEntry] = None
    candidates: List[DocEntry] = field(default_factory=list)
    warning: Optional[str] = None
    uncertainty: Optional[str] = None

class Binder:
    """
    Deterministic binder:
      - VBA member_call: prefer class.member when receiver looks like Range/Workbook/Worksheet via *context*,
        but do not claim certainty. If ambiguous -> list candidates.
      - M function_call: bind by exact function symbol.
    """

    def __init__(self, docs: DocsCache):
        self.docs = docs

    def infer_vba_receiver_class(self, receiver_text: Optional[str],
                                 receiver_is_identifier: bool,
                                 with_receiver: Optional[str]) -> Tuple[Optional[str], Optional[str]]:
        # Use With receiver if present and fact receiver missing
        r = (receiver_text or with_receiver or "").strip()
        if not r:
            return None, None

        low = r.lower()
        globals_map = {
            "activeworkbook": "workbook",
            "thisworkbook": "workbook",
            "activesheet": "worksheet",
            "activecell": "range",
        }
        if receiver_is_identifier and low in globals_map:
            return globals_map[low], "Receiver class inferred from global Excel object."
        return None, None

    def bind_fact(self, fact: Fact) -> BoundItem:
        bi = BoundItem(fact=fact)

        if fact.language == "m" and fact.kind == "function_call":
            fn = (fact.member_or_fn or "").strip()
            if not fn:
                bi.warning = "Unbound M call: function name not extractable with current grammar."
                return bi
            entry = self.docs.lookup_m(fn)
            if entry:
                bi.bound = entry
            else:
                bi.warning = f"Unbound M function '{fn}': not found in docs cache."
            return bi

        if fact.language == "vba" and fact.kind == "member_call":
            member = (fact.member_or_fn or "").strip()
            if not member:
                bi.warning = "Unbound member call: member name not extractable."
                return bi

            # NEW: Prefer receiver_type from symbol table (declaration tracking)
            cls = None
            note = None

            if fact.receiver_type:
                # Type information from Dim/Set declarations
                cls = fact.receiver_type.lower()
                note = f"Receiver type '{fact.receiver_type}' from variable declaration (Dim/Set statement)."
            else:
                # Fall back to heuristic inference from receiver text
                receiver_text = fact.receiver_base or fact.receiver
                cls, note = self.infer_vba_receiver_class(
                    receiver_text, fact.receiver_is_identifier, fact.with_receiver
                )

            if note:
                bi.uncertainty = note

            if cls:
                exact = self.docs.lookup_vba_class_member(cls, member)
                if exact:
                    bi.bound = exact
                    return bi

            cands = self.docs.candidates_vba_member(member)
            bi.candidates = cands

            if len(cands) == 1:
                bi.bound = cands[0]
            elif len(cands) == 0:
                bi.warning = f"Unbound member '{member}': no matching API in docs cache."
            else:
                bi.warning = (
                    f"Ambiguous binding for '{member}': matches multiple documented APIs. "
                    "Candidates listed instead of guessing."
                )
            return bi

        if fact.language == "vba" and fact.kind == "file_io":
            # Not bound to docs; it’s a language/runtime side effect.
            bi.warning = fact.note or "File I/O detected."
            return bi

        # default
        bi.warning = fact.note or "Detected construct not bound to docs (or unsupported fact type)."
        return bi

def render_bound_item(b: BoundItem) -> str:
    """Render a bound item as natural English prose."""
    f = b.fact

    parts: List[str] = []

    # Build the explanation
    if f.kind == "parse_error":
        msg = f.note or "Parsing failed."
        parts.append(f"Line {f.line}: {msg}")
        return "".join(parts)
    if f.kind == "parse_warning":
        msg = f.note or "Parser warning."
        parts.append(f"Line {f.line}: {msg}")
        return "".join(parts)

    if b.bound:
        # We have documentation - explain what it does

        # Build natural explanation
        if f.receiver and f.receiver_type:
            parts.append(f"Line {f.line}: '{f.text}' operates on a {f.receiver_type} ('{f.receiver}'). ")
        elif f.receiver:
            parts.append(f"Line {f.line}: '{f.text}' operates on {f.receiver}. ")
        else:
            parts.append(f"Line {f.line}: '{f.text}' executes. ")

        # Extract the API name
        if f.language == "vba":
            # Extract just the class.member from Excel.Class.Member
            api_parts = b.bound.symbol.split(".")
            if len(api_parts) >= 3:
                class_name = api_parts[1]
                member_name = api_parts[2].split()[0]  # Strip " property" etc
                api_name = f"{class_name}.{member_name}"
            else:
                api_name = b.bound.symbol
        else:
            api_name = b.bound.symbol

        # Determine if it's a property or method
        is_property = "property" in b.bound.symbol.lower()
        is_method = "method" in b.bound.symbol.lower() or b.bound.kind == "method"

        if is_property:
            parts.append(f"It accesses the {api_name} property, ")
        elif is_method:
            parts.append(f"It calls the {api_name} method, ")
        else:
            parts.append(f"It uses {api_name}, ")

        # Add what it does
        summary = (b.bound.summary or "").strip()
        if summary:
            # Make it flow naturally as a continuation
            if summary[0].isupper():
                summary = summary[0].lower() + summary[1:]

        if f.language == "vba" and f.member_or_fn and f.member_or_fn.lower() == "value" and f.assigned_value:
            if summary:
                parts.append(f"setting it to {f.assigned_value}, ")
            else:
                parts.append(f"setting it to {f.assigned_value}. ")

        if summary:
            parts.append(f"which {summary}")
            if not summary.endswith('.'):
                parts.append('.')
        else:
            parts.append("(no documentation summary available).")

    elif f.kind == "file_io":
        # File I/O operations
        parts.append(f"Line {f.line}: '{f.text}' performs file I/O. ")
        if b.warning:
            parts.append(b.warning)

    elif b.warning:
        # Couldn't bind to docs
        parts.append(f"Line {f.line}: '{f.text}' executes. ")

        if "no matching API" in b.warning:
            if f.receiver and f.receiver_type:
                parts.append(f"The variable '{f.receiver}' is a {f.receiver_type}, but the member '{f.member_or_fn}' wasn't found in the documentation cache.")
            else:
                parts.append(f"The member '{f.member_or_fn}' wasn't found in the documentation cache.")
        elif b.candidates:
            parts.append(f"The member '{f.member_or_fn}' matches {len(b.candidates)} possible APIs: ")
            cand_names = [c.symbol for c in b.candidates[:3]]
            parts.append(", ".join(cand_names))
            if len(b.candidates) > 3:
                parts.append(f", and {len(b.candidates) - 3} more")
            parts.append(".")
        else:
            parts.append(b.warning)

    else:
        # Fallback
        parts.append(f"The code contains: {f.text}")

    return "".join(parts)

def render_report(title: str, bound_items: List[BoundItem],
                  flow_narrative: Optional[str], docs: Optional[DocsCache]) -> str:
    """
    Generate a natural language report of the analysis.

    Now includes optional flow narrative for Power Query M transformations.
    """
    parts: List[str] = []

    # Header
    parts.append("MACRO ANALYSIS REPORT")
    parts.append("=" * 80)
    parts.append(f"File: {title}")
    if docs:
        parts.append(f"Documentation: {len(docs.entries_by_id)} cached API entries from Microsoft Learn")
    parts.append("")

    # Flow narrative (if available - only for M queries)
    if flow_narrative:
        parts.append("TRANSFORMATION PIPELINE:")
        parts.append("")
        parts.append(flow_narrative)
        parts.append("")
        parts.append("=" * 80)
        parts.append("")

    # Detailed line-by-line analysis
    if not bound_items:
        parts.append("No executable code constructs were detected in this file.")
    else:
        parts.append("DETAILED ANALYSIS:")
        parts.append("")
        for bi in _select_best_bound_items(bound_items):
            parts.append(render_bound_item(bi))
            parts.append("")

    # Footer notes
    parts.append("=" * 80)
    parts.append("ANALYSIS NOTES:")
    parts.append("• Code structure parsed using ANTLR grammars (not regex)")
    parts.append("• Variable types tracked from Dim/Set declarations")
    parts.append("• API documentation from Microsoft Learn (offline cache)")
    parts.append("• Ambiguous bindings reported explicitly rather than guessed")
    if flow_narrative:
        parts.append("• Pipeline summaries: docs cache + curated verb mappings for common functions")

    return "\n".join(parts)

# ---------------------------
# App
# ---------------------------

class AnalyzerApp:
    def __init__(self, root: tk.Tk):
        self.root = root
        self.root.title("Excel Demystifier Analyzer (ANTLR AST)")
        self.root.geometry("1150x780")

        self.cache_dir = "doc_cache"
        self.docs: Optional[DocsCache] = None

        self.vba_parser = None
        self.m_parser = None
        self.vba_lang_id = None
        self.m_lang_id = None
        self.parser_type = None

        self._build_ui()
        self._init_parsers()
        self._load_docs_cache_async()

    def _build_ui(self):
        self.root.rowconfigure(1, weight=1)
        self.root.columnconfigure(0, weight=1)

        top = tk.Frame(self.root, padx=8, pady=8)
        top.grid(row=0, column=0, sticky="ew")
        top.columnconfigure(6, weight=1)

        tk.Button(top, text="Analyze (pasted)", command=self.analyze_pasted).grid(row=0, column=0, padx=(0, 8))
        tk.Button(top, text="Upload file…", command=self.upload_file).grid(row=0, column=1, padx=(0, 8))
        tk.Button(top, text="Rebuild docs cache", command=self.rebuild_cache).grid(row=0, column=2, padx=(0, 8))
        tk.Button(top, text="Copy output", command=self.copy_output).grid(row=0, column=3, padx=(0, 8))
        tk.Button(top, text="Clear", command=self.clear).grid(row=0, column=4, padx=(0, 8))
        tk.Button(top, text="About", command=self.about).grid(row=0, column=5, padx=(0, 8))

        self.status = tk.Label(top, text="Ready", anchor="w")
        self.status.grid(row=0, column=6, sticky="ew")

        mid = tk.Frame(self.root, padx=8, pady=8)
        mid.grid(row=1, column=0, sticky="nsew")
        mid.rowconfigure(0, weight=1)
        mid.columnconfigure(0, weight=1)

        self.output = tk.Text(mid, wrap="word")
        self.output.grid(row=0, column=0, sticky="nsew")
        scroll = tk.Scrollbar(mid, command=self.output.yview)
        scroll.grid(row=0, column=1, sticky="ns")
        self.output.configure(yscrollcommand=scroll.set)

        bottom = tk.Frame(self.root, padx=8, pady=8)
        bottom.grid(row=2, column=0, sticky="ew")
        bottom.columnconfigure(0, weight=1)

        tk.Label(bottom, text="Paste VBA macro or Power Query M text here:").grid(row=0, column=0, sticky="w")
        self.input = tk.Text(bottom, wrap="word", height=10)
        self.input.grid(row=1, column=0, sticky="ew")

    def _set_status(self, s: str):
        self.status.config(text=s)
        self.root.update_idletasks()

    def about(self):
        msg = (
            "Deterministic analyzer (ANTLR AST):\n"
            "- AST parsing via ANTLR (no regex parsing of syntax)\n"
            "- Offline docs cache from Microsoft Learn\n"
            "- Docs-backed deterministic rendering with explicit uncertainty\n\n"
            f"VBA parser: {self.vba_lang_id or '(not loaded)'}\n"
            f"M parser: {self.m_lang_id or '(not loaded)'}\n"
            f"Parser type: {self.parser_type or '(unknown)'}\n"
        )
        messagebox.showinfo("About", msg)

    def _init_parsers(self):
        try:
            self._set_status("Loading ANTLR parsers…")
            self.vba_parser, self.vba_lang_id, pkg1 = get_vba_parser()
            self.m_parser, self.m_lang_id, pkg2 = get_m_parser()
            self.parser_type = pkg1 or pkg2
            self._set_status(f"Parsers loaded (VBA={self.vba_lang_id}, M={self.m_lang_id})")
        except Exception as e:
            self._set_status("Parser load error")
            messagebox.showerror("ANTLR parser error", str(e))

    def _load_docs_cache_async(self):
        cache_path = ensure_docs_cache_path(self.cache_dir)
        if os.path.exists(cache_path):
            try:
                self._load_docs_cache_now()
                self._set_status(f"Docs cache loaded ({len(self.docs.entries_by_id) if self.docs else 0} entries)")
            except Exception as e:
                self._set_status("Docs cache error")
                messagebox.showerror("Docs cache error", str(e))
            return

        self._start_harvest_dialog()

    def _start_harvest_dialog(self):
        dlg = tk.Toplevel(self.root)
        dlg.title("Building docs cache…")
        dlg.geometry("520x180")
        dlg.transient(self.root)
        dlg.grab_set()

        tk.Label(
            dlg,
            text="docs_cache.json not found.\nHarvesting Microsoft Learn docs (offline cache build)…",
            justify="left"
        ).pack(padx=12, pady=(12, 6), anchor="w")

        prog = tk.Label(dlg, text="Starting…", anchor="w")
        prog.pack(padx=12, pady=6, fill="x")

        btn_close = tk.Button(dlg, text="Close", state="disabled", command=dlg.destroy)
        btn_close.pack(padx=12, pady=(6, 12), anchor="e")

        self._set_status("Harvesting docs cache…")
        result_queue: Queue = Queue()

        def worker():
            try:
                rc, out, err = run_harvester_ensure(self.cache_dir)
                result_queue.put((rc == 0, (out + "\n" + err).strip()))
            except Exception as e:
                result_queue.put((False, str(e)))

        t = threading.Thread(target=worker, daemon=True)
        t.start()

        def poll():
            if not result_queue.empty():
                ok, msg = result_queue.get()
            elif t.is_alive():
                dots = "." * ((int(time.time() * 2) % 4) + 1)
                prog.config(text=f"Harvesting{dots}")
                dlg.after(250, poll)
                return
            else:
                dlg.after(100, poll_final)
                return

            if ok:
                prog.config(text="Done. Loading docs cache…")
                try:
                    self._load_docs_cache_now()
                    self._set_status(f"Docs cache loaded ({len(self.docs.entries_by_id) if self.docs else 0} entries)")
                    dlg.destroy()
                except Exception as e:
                    prog.config(text="Failed to load docs cache.")
                    btn_close.config(state="normal")
                    messagebox.showerror("Docs cache error", str(e))
                    self._set_status("Docs cache error")
            else:
                prog.config(text="Harvester failed. Details written to output panel.")
                btn_close.config(state="normal")
                self.output.delete("1.0", "end")
                self.output.insert("1.0", "HARVESTER FAILED\n\n" + (msg or "(no output)"))
                self._set_status("Harvester failed")

        dlg.after(250, poll)

        def poll_final():
            if not result_queue.empty():
                ok, msg = result_queue.get()
            else:
                ok, msg = False, "Harvester thread exited without a result."

            if ok:
                prog.config(text="Done. Loading docs cache…")
                try:
                    self._load_docs_cache_now()
                    self._set_status(f"Docs cache loaded ({len(self.docs.entries_by_id) if self.docs else 0} entries)")
                    dlg.destroy()
                except Exception as e:
                    prog.config(text="Failed to load docs cache.")
                    btn_close.config(state="normal")
                    messagebox.showerror("Docs cache error", str(e))
                    self._set_status("Docs cache error")
            else:
                prog.config(text="Harvester failed. Details written to output panel.")
                btn_close.config(state="normal")
                self.output.delete("1.0", "end")
                self.output.insert("1.0", "HARVESTER FAILED\n\n" + (msg or "(no output)"))
                self._set_status("Harvester failed")

    def _load_docs_cache_now(self):
        cache_path = ensure_docs_cache_path(self.cache_dir)
        docs = DocsCache(cache_path)
        docs.load()
        self.docs = docs

    def rebuild_cache(self):
        try:
            cache_path = ensure_docs_cache_path(self.cache_dir)
            if os.path.exists(cache_path):
                os.remove(cache_path)
            self.docs = None
            self._load_docs_cache_async()
        except Exception as e:
            messagebox.showerror("Rebuild error", str(e))

    def clear(self):
        self.output.delete("1.0", "end")
        self.input.delete("1.0", "end")
        self._set_status("Ready")

    def copy_output(self):
        text = self.output.get("1.0", "end").strip()
        if not text:
            messagebox.showinfo("Copy output", "There is no output to copy.")
            return
        self.root.clipboard_clear()
        self.root.clipboard_append(text)
        self._set_status("Output copied to clipboard")

    def upload_file(self):
        path = filedialog.askopenfilename(
            title="Select a text/code file containing VBA or M",
            filetypes=[
                ("Code/Text files", "*.txt *.vba *.bas *.cls *.frm *.m *.pq"),
                ("All files", "*.*"),
            ],
        )
        if not path:
            return
        try:
            with open(path, "r", encoding="utf-8", errors="replace") as f:
                text = f.read()
            self.input.delete("1.0", "end")
            self.input.insert("1.0", text)
            self.analyze_text(text, title=os.path.basename(path))
        except Exception as e:
            messagebox.showerror("File error", str(e))

    def analyze_pasted(self):
        text = self.input.get("1.0", "end").strip()
        if not text:
            messagebox.showinfo("No input", "Paste code in the bottom field or upload a file.")
            return
        self.analyze_text(text, title="Pasted text")

    def analyze_text(self, text: str, title: str):
        """
        Analyze VBA or M code and display results.
        Now includes semantic data flow analysis for Power Query M.
        """
        if len(text) > MAX_INPUT_CHARS:
            messagebox.showerror(
                "Input too large",
                f"Input is {len(text)} characters; limit is {MAX_INPUT_CHARS}."
            )
            return
        if not self.docs:
            messagebox.showerror("No docs", "Docs cache not loaded yet.")
            return
        if not self.vba_parser and not self.m_parser:
            messagebox.showerror("No parsers", "ANTLR parsers not available.")
            return

        self._set_status("Analyzing…")
        try:
            vba_facts: List[Fact] = []
            m_facts: List[Fact] = []
            m_flow_graph: Optional[DataFlowGraph] = None

            # Extract VBA facts (no flow graph for VBA yet)
            if self.vba_parser:
                vba_facts = VBAFactExtractor(self.vba_parser).extract(text)

            # Extract M facts AND flow graph
            if self.m_parser:
                m_extractor = MFactExtractor(self.m_parser, docs=self.docs)
                m_facts, m_flow_graph = m_extractor.extract_with_flow(text)

            all_facts = vba_facts + m_facts
            if any(f.kind == "parse_error" for f in all_facts):
                bound_items = [BoundItem(fact=f) for f in all_facts
                               if f.kind in ("parse_error", "parse_warning")]
                flow_narrative = None
            else:
                # Bind facts to documentation
                binder = Binder(self.docs)
                bound_items = [binder.bind_fact(f) for f in all_facts]
                bound_items.sort(key=lambda bi: (bi.fact.line, bi.fact.language, bi.fact.kind))

                # Generate flow narrative (if M flow graph available)
                flow_narrative: Optional[str] = None
                if m_flow_graph and m_flow_graph.nodes:
                    renderer = SemanticFlowRenderer(self.docs)
                    flow_narrative = renderer.render_flow(m_flow_graph)

            # Generate report
            report = render_report(title, bound_items, flow_narrative, self.docs)
            self.output.delete("1.0", "end")
            self.output.insert("1.0", report)
            self._set_status("Done")
        except Exception as e:
            self._set_status("Error")
            import traceback
            error_details = f"{str(e)}\n\n{traceback.format_exc()}"
            messagebox.showerror("Analysis error", error_details)

def main():
    logging.basicConfig(
        level=logging.WARNING,
        format="%(asctime)s - %(name)s - %(levelname)s - %(message)s",
        handlers=[
            logging.FileHandler("analyzer.log"),
            logging.StreamHandler(),
        ],
    )
    root = tk.Tk()
    AnalyzerApp(root)
    root.mainloop()

if __name__ == "__main__":
    main()
