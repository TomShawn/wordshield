#!/bin/zsh
set -euo pipefail
exec "$(dirname "$0")/scripts/dev.sh" run
