#!/bin/sh
# Recorde — sideload the Word add-in for the current macOS user.
#
# Copies manifest.xml (next to this script) into Word's sideload folder,
# ~/Library/Containers/com.microsoft.Word/Data/Documents/wef, as Microsoft's
# "Sideload Office Add-ins on Mac" procedure describes. No admin rights.
# Undo with uninstall-mac.command.
#
# Double-clicking a downloaded .command file may be blocked by Gatekeeper;
# then run it from Terminal:  sh install-mac.command
set -e
DIR="$(cd "$(dirname "$0")" && pwd)"
WEF="$HOME/Library/Containers/com.microsoft.Word/Data/Documents/wef"

if [ ! -f "$DIR/manifest.xml" ]; then
  echo "manifest.xml was not found next to this script. Extract the whole zip first."
  exit 1
fi
mkdir -p "$WEF"
cp "$DIR/manifest.xml" "$WEF/recorde.manifest.xml"

echo "Recorde is sideloaded: $WEF/recorde.manifest.xml"
echo "Quit Word completely and start it again. Home > Add-ins lists Recorde (under Developer Add-ins)."
