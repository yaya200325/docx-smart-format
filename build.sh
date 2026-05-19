#!/usr/bin/env bash
# Build the docx-auto-template-engine and publish a self-contained runtime.
#
# Daily skill users do NOT need this — only run when you change C# source
# or need a non-default RID build.
#
# Usage:
#   ./build.sh                    # defaults: --rid linux-x64 -c Release
#   ./build.sh --rid osx-arm64
#   ./build.sh --rid win-x64 -c Debug

set -euo pipefail

RID="linux-x64"
CONFIGURATION="Release"
OUTPUT_DIR="engine/runtime"

print_usage() {
  cat <<'EOF'
Usage: ./build.sh [options]

Options:
  --rid <id>            Runtime identifier (default: linux-x64). Examples:
                        linux-x64, linux-arm64, osx-arm64, osx-x64, win-x64.
  -c, --configuration   Build configuration (default: Release).
  -o, --output          Output directory (default: engine/runtime).
  -h, --help            Show this help and exit.
EOF
}

while [[ $# -gt 0 ]]; do
  case "$1" in
    --rid)
      RID="$2"; shift 2 ;;
    --rid=*)
      RID="${1#*=}"; shift ;;
    -c|--configuration)
      CONFIGURATION="$2"; shift 2 ;;
    --configuration=*)
      CONFIGURATION="${1#*=}"; shift ;;
    -o|--output)
      OUTPUT_DIR="$2"; shift 2 ;;
    --output=*)
      OUTPUT_DIR="${1#*=}"; shift ;;
    -h|--help)
      print_usage; exit 0 ;;
    *)
      echo "Unknown option: $1" >&2
      print_usage; exit 1 ;;
  esac
done

SCRIPT_DIR="$( cd "$( dirname "${BASH_SOURCE[0]}" )" && pwd )"
PROJECT_PATH="${SCRIPT_DIR}/engine/src/docx-auto-template-engine.csproj"

if [[ ! -f "$PROJECT_PATH" ]]; then
  echo "Project file not found at $PROJECT_PATH" >&2
  echo "Run this script from the repository root." >&2
  exit 1
fi

if ! command -v dotnet >/dev/null 2>&1; then
  echo ".NET SDK not found. Install .NET SDK 8.0+ from https://dotnet.microsoft.com/download" >&2
  exit 1
fi

echo "Building docx-auto-template-engine"
echo "  RID:           $RID"
echo "  Configuration: $CONFIGURATION"
echo "  Output:        $OUTPUT_DIR"
echo

dotnet publish "$PROJECT_PATH" \
  -c "$CONFIGURATION" \
  -r "$RID" \
  --self-contained true \
  -p:PublishSingleFile=false \
  -o "${SCRIPT_DIR}/${OUTPUT_DIR}"

echo
echo "Build complete. Engine published to ${SCRIPT_DIR}/${OUTPUT_DIR}"
