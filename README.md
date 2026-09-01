# Excel Demystifier Analyzer

A desktop tool that explains what Excel VBA macros and Power Query M scripts
do — without ever executing them.

Paste (or upload) macro code you inherited or don't trust, and it produces a
plain-language report of the code's structure and effects: which workbooks,
sheets and ranges it touches, what it writes, sorts, deletes or saves, and
what each API call means according to Microsoft's own documentation.

## How it works

- **Static analysis only.** The code is parsed, never run. This makes it safe
  to point at untrusted macros — a useful first step before enabling anything
  in Excel.
- **Real parsers, not regex.** Source text is parsed into an AST with ANTLR
  grammars (VBA and Power Query M). Facts are extracted deterministically from
  the tree — there is no heuristic pattern-matching of syntax.
- **Docs-backed semantics.** The meaning of each call is looked up in an
  offline cache of Microsoft Learn reference pages (built once by
  `harvest_docs.py`). If the analyzer can't ground a claim in the grammar or
  the docs, it says it is uncertain rather than guessing.

Two programs:

| File | Purpose |
|------|---------|
| `analyzer.py` | Tkinter GUI: parse VBA / Power Query M, extract facts from the AST, bind them to cached Microsoft Learn docs, render a report |
| `harvest_docs.py` | Scope-bounded crawler that builds the offline docs cache (`doc_cache/`) from the Excel VBA API and Power Query M reference on learn.microsoft.com |

## Setup

Requires Python 3.10+ and, only for regenerating the parsers, a Java runtime.

```sh
pip install -r requirements.txt
```

### 1. Get the grammars

The VBA grammar is vendored in `grammars/vba/` (MIT, see below). The Power
Query M grammar is fetched, because its upstream repository does not include
a license file that would permit redistributing it here:

```sh
./fetch_grammars.sh
```

This downloads the two Power Query grammar files from
[antlr/grammars-v4](https://github.com/antlr/grammars-v4) at a pinned commit
and applies `grammars/powerquery/local-fixes.patch` — this project's own
grammar fixes (case-insensitive lexing, manual left-recursion elimination,
and an `item_selector` rule that breaks a mutual recursion).

### 2. Generate the parsers

Generated parser code and the ANTLR tool jar (~2 MB binary) are not
committed. With Java installed:

```sh
./generate_parsers.sh
```

This downloads `antlr-4.13.2-complete.jar` from
<https://www.antlr.org/download/antlr-4.13.2-complete.jar> if it is not
already present, generates Python parsers into `parsers/vba/` and
`parsers/powerquery/`, and applies one small fix: the upstream VBA grammar
names two tokens `True` and `False`, which are Python keywords, so ANTLR's
Python3 target emits invalid syntax for them — the script renames those
references so the module imports.

The ANTLR tool version (4.13.2) must match `antlr4-python3-runtime` in
`requirements.txt`.

### 3. Build the docs cache (optional but recommended)

```sh
python3 harvest_docs.py harvest --out-dir doc_cache
```

This crawls only the Excel VBA API and Power Query M sections of
learn.microsoft.com (throttled, cached, domain- and path-bounded) and writes
`doc_cache/docs_cache.json`. It takes a while and produces a few hundred MB.
The analyzer also offers to build the cache on first launch; without a cache
it still parses and reports structure, just without doc-backed descriptions.

`python3 harvest_docs.py validate --out-dir doc_cache` checks extraction
quality afterwards — it warns if Microsoft's page layout has drifted enough
to degrade the cache.

## Usage

```sh
python3 analyzer.py
```

Paste code into the input box (or use *Upload file…*) and click *Analyze*.
Two small inputs are included to try it on:

- `test_vba.vba` — a macro that writes to and sorts ranges, then saves the
  workbook
- `test_m.pq` — a Power Query `let` expression that filters, sorts and
  removes columns from a table

## Grammar attribution

- **VBA**: [rossknudsen/Vba.Language](https://github.com/rossknudsen/Vba.Language),
  MIT licensed. The four `.g4` files are vendored unmodified in
  `grammars/vba/` together with their upstream `LICENSE`.
- **Power Query M**: [antlr/grammars-v4](https://github.com/antlr/grammars-v4)
  (`powerquery/`). That grammar ships without a license file, so it is not
  vendored; `fetch_grammars.sh` downloads it from upstream and applies this
  project's fixes on top.

## License

MIT — see [LICENSE](LICENSE). The vendored VBA grammar retains its own MIT
license in `grammars/vba/LICENSE`.
