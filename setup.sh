#!/bin/bash
# Eden NewGen — Automated Setup Script
# Usage: ./setup.sh

set -e

echo "🌻 Eden NewGen — Setup"
echo "========================"
echo ""

# Check Python
if ! command -v python3 &> /dev/null; then
    echo "❌ Python 3 not found. Install: apt install python3 python3-pip"
    exit 1
fi
echo "✓ Python: $(python3 --version)"

# Check Git
if ! command -v git &> /dev/null; then
    echo "❌ Git not found. Install: apt install git"
    exit 1
fi
echo "✓ Git: $(git --version)"

# Step 1: Clone web2py
EDEN_DIR="$(pwd)/study_repo/eden"
WEB2PY_DIR="$(pwd)/web2py"

if [ ! -d "$WEB2PY_DIR" ]; then
    echo ""
    echo "=== Step 1: Clone web2py ==="
    git clone https://github.com/web2py/web2py.git "$WEB2PY_DIR"
    cd "$WEB2PY_DIR"
    git reset --hard 9005306
    git submodule update --init --recursive
    cd ..
    echo "✓ web2py cloned"
else
    echo "✓ web2py already exists"
fi

# Symlink Eden into web2py
if [ ! -L "$WEB2PY_DIR/applications/eden" ]; then
    ln -s "$EDEN_DIR" "$WEB2PY_DIR/applications/eden"
    echo "✓ Eden symlinked to web2py/applications/eden"
fi

# Step 2: Install Python dependencies
echo ""
echo "=== Step 2: Install Python dependencies ==="
pip install -r "$EDEN_DIR/requirements.txt" 2>/dev/null || true
pip install -r "$EDEN_DIR/optional_requirements.txt" 2>/dev/null || true
pip install -r "$EDEN_DIR/fastapi_app/requirements.txt" 2>/dev/null || true
pip install structlog fastapi uvicorn sqlalchemy aiosqlite pydantic pydantic-settings passlib pyotp qrcode 2>/dev/null || true
echo "✓ Dependencies installed"

# Step 3: Configure Eden
echo ""
echo "=== Step 3: Configure Eden ==="
if [ ! -f "$EDEN_DIR/models/000_config.py" ]; then
    cp "$EDEN_DIR/modules/templates/000_config.py" "$EDEN_DIR/models/000_config.py"
    sed -i 's/FINISHED_EDITING_CONFIG_FILE = False/FINISHED_EDITING_CONFIG_FILE = True/' "$EDEN_DIR/models/000_config.py"
    echo "✓ Config created"
else
    echo "✓ Config already exists"
fi

# Step 4: First-time setup
echo ""
echo "=== Step 4: First-time setup (database creation) ==="
if [ ! -f "$EDEN_DIR/databases/storage.db" ]; then
    cd "$WEB2PY_DIR"
    python3 web2py.py -S eden -M -R applications/eden/static/scripts/tools/noop.py 2>&1 | tail -5
    echo "✓ Database created"
else
    echo "✓ Database already exists"
fi

# Step 5: Create .env for FastAPI
echo ""
echo "=== Step 5: Create .env ==="
if [ ! -f "$EDEN_DIR/.env" ]; then
    cp "$EDEN_DIR/.env.example" "$EDEN_DIR/.env"
    echo "✓ .env created (edit with real API keys)"
else
    echo "✓ .env already exists"
fi

echo ""
echo "========================"
echo "✅ Setup complete!"
echo ""
echo "To start:"
echo "  # web2py:"
echo "  cd $WEB2PY_DIR"
echo "  python3 web2py.py --no_gui -a admin --port 8000"
echo ""
echo "  # FastAPI:"
echo "  cd $EDEN_DIR/fastapi_app"
echo "  PYTHONPATH=. uvicorn main:app --reload --port 8001"
echo ""
echo "  # Next.js:"
echo "  cd $(pwd)"
echo "  npm install && npm run dev"
echo ""
echo "  Login: admin@example.com / admin123"
echo "========================"
