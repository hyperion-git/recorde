#!/bin/sh
# Recorde — remove the sideloaded Word add-in manifest (install-mac.command).
WEF="$HOME/Library/Containers/com.microsoft.Word/Data/Documents/wef"
rm -f "$WEF/recorde.manifest.xml"
echo "Recorde is removed from $WEF. Quit Word completely and start it again."
echo "If the button lingers, clear the Office cache (see INSTALL.md - Updating / Troubleshooting)."
