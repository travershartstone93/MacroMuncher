#!/usr/bin/env bash
# Fetch the Power Query M grammar from antlr/grammars-v4 and apply local fixes.
#
# The Power Query grammar in antlr/grammars-v4 does not ship a license file,
# so it is not vendored in this repository. This script downloads the two
# grammar files at a pinned commit (the revision this project was built
# against, before the 2023 repo-wide reformat) and applies
# grammars/powerquery/local-fixes.patch (this project's own modifications:
# case-insensitive lexing, manual left-recursion elimination, and an
# item_selector fix that breaks a mutual recursion).
#
# The VBA grammar (grammars/vba/) is vendored directly because its upstream,
# github.com/rossknudsen/Vba.Language, is MIT licensed.
set -euo pipefail

cd "$(dirname "$0")"

COMMIT="4bfa06b0188e129aa390159a3a209385aed00014"
BASE="https://raw.githubusercontent.com/antlr/grammars-v4/${COMMIT}/powerquery"
DEST="grammars/powerquery"

curl -fsSL "${BASE}/PowerQueryLexer.g4" -o "${DEST}/PowerQueryLexer.g4"
curl -fsSL "${BASE}/PowerQueryParser.g4" -o "${DEST}/PowerQueryParser.g4"

sha256sum -c --quiet <<EOF
a410857b325aeb4be5e071d0471cb579b23ccd0afc90737542317fa75039d811  ${DEST}/PowerQueryLexer.g4
d55d5127613e5ed6abff07c3789a03d527290a67dfd10749c0fcdf27cadab690  ${DEST}/PowerQueryParser.g4
EOF

patch -p1 -d "${DEST}" < "${DEST}/local-fixes.patch"

echo "Power Query grammar fetched and patched in ${DEST}/"
