#!/bin/bash
cd ~/my-web-apps
git add .
git commit -m "update: $(date '+%Y-%m-%d %H:%M')"
git push
echo ""
echo "✓ Pushed to GitHub. Netlify will go live in ~30 seconds."
