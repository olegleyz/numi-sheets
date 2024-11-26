#!/bin/bash

# Check if clasp is installed
if ! command -v clasp &> /dev/null; then
    echo "clasp is not installed. Installing clasp..."
    npm install -g @google/clasp
fi

# Only login if not already logged in
if [ ! -f ~/.clasprc.json ]; then
    echo "Not logged in. Logging in to clasp..."
    clasp login
fi

# Only create a new project if .clasp.json doesn't exist
if [ ! -f .clasp.json ] && [ ! -f src/.clasp.json ]; then
    echo "Creating new Google Apps Script project..."
    clasp create --type sheets --title "Numi Foreseer" --rootDir ./src
elif [ -f src/.clasp.json ] && [ ! -f .clasp.json ]; then
    echo "Moving .clasp.json from src to root directory..."
    mv src/.clasp.json .
fi

# Push changes to Google Apps Script
echo "Deploying changes..."
clasp push

echo "Deployment complete!"
