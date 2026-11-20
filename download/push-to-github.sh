#!/bin/bash
# Eden NewGen — Push to GitHub
# Run this on YOUR machine (not in the sandbox)

set -e

echo "=== Eden NewGen → GitHub ==="

# Option 1: If you downloaded eden-newgen.tar.gz
if [ -f "eden-newgen.tar.gz" ]; then
    tar xzf eden-newgen.tar.gz
    cd eden-newgen
fi

# Option 2: If you downloaded eden-newgen.bundle
if [ -f "eden-newgen.bundle" ]; then
    git clone eden-newgen.bundle eden-newgen
    cd eden-newgen
    git remote remove origin 2>/dev/null || true
fi

# Initialize if needed
git init 2>/dev/null || true
git add -A
git commit -m "Eden NewGen: FastAPI + AI + Next.js + Docker + Helm" 2>/dev/null || true
git branch -M main 2>/dev/null || true

# Add remote
git remote add origin https://github.com/skpuan07-cmd/eden-newgen.git 2>/dev/null || \
git remote set-url origin https://github.com/skpuan07-cmd/eden-newgen.git

# Push!
echo "Pushing to GitHub..."
git push -u origin main

echo ""
echo "✅ Done! https://github.com/skpuan07-cmd/eden-newgen"
