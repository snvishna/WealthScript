#!/bin/bash
# Run this from the root of the repository
cat src/Config.gs src/Formulas.gs src/Menu.gs src/Wizard.gs src/API.gs src/Builders.gs src/Dashboards.gs src/Snapshot.gs src/Backup.gs src/Repair.gs > deploy/code.gs
cp appsscript.json deploy/appsscript.json

echo "build.sh: Successfully generated deploy/code.gs and deploy/appsscript.json"
