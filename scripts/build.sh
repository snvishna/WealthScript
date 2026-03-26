#!/bin/bash
# Run this from the root of the repository
cat src/Config.gs src/Menu.gs src/API.gs src/Builders.gs src/Dashboards.gs src/Snapshot.gs src/Backup.gs > deploy/code.gs

echo "build.sh: Successfully generated deploy/code.gs"
