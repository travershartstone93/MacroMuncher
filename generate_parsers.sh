#!/usr/bin/env bash
# Generate the Python parsers from the ANTLR grammars into parsers/.
#
# Requires: java, and the ANTLR tool jar (downloaded automatically if absent).
# The tool version must match antlr4-python3-runtime in requirements.txt.
#
# Post-generation fix: the upstream VBA lexer grammar names two tokens "True"
# and "False", which are Python keywords, so ANTLR's Python3 target emits
# invalid syntax for them in VbaParser.py (e.g. "VbaParser.True"). The sed
# below renames those two references so the module imports cleanly.
set -euo pipefail

cd "$(dirname "$0")"

JAR="antlr-4.13.2-complete.jar"
if [ ! -f "$JAR" ]; then
    curl -fsSLO "https://www.antlr.org/download/${JAR}"
fi

if [ ! -f grammars/powerquery/PowerQueryLexer.g4 ]; then
    ./fetch_grammars.sh
fi

mkdir -p parsers/vba parsers/powerquery

(cd grammars/vba && java -jar "../../${JAR}" -Dlanguage=Python3 -o ../../parsers/vba VbaLexer.g4 VbaParser.g4)
(cd grammars/powerquery && java -jar "../../${JAR}" -Dlanguage=Python3 -o ../../parsers/powerquery PowerQueryLexer.g4 PowerQueryParser.g4)

sed -i 's/VbaParser\.True\b/VbaParser.TRUE_/g; s/VbaParser\.False\b/VbaParser.FALSE_/g' parsers/vba/VbaParser.py

python3 -m py_compile parsers/vba/VbaLexer.py parsers/vba/VbaParser.py \
    parsers/powerquery/PowerQueryLexer.py parsers/powerquery/PowerQueryParser.py

echo "Parsers generated in parsers/"
