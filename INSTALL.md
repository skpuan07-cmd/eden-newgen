# วิธีติดตั้ง Eden NewGen (ฉบับสมบูรณ์)

> คู่มือติดตั้งทีละขั้นตอน — ทำตามได้บนเครื่อง Linux/Mac ของคุณ

---

## ขั้นที่ 1: Clone repository

```bash
git clone https://github.com/skpuan07-cmd/eden-newgen.git
cd eden-newgen
```

## ขั้นที่ 2: ติดตั้ง web2py framework

Eden ทำงานบน web2py จึงต้อง clone web2py ก่อน:

```bash
cd /path/to/eden-newgen/study_repo/eden

# Clone web2py (รุ่น stable 3.3.3)
git clone https://github.com/web2py/web2py.git ../../web2py
cd ../../web2py
git reset --hard 9005306
git submodule update --init --recursive

# Symlink Eden เข้า web2py applications
cd applications
ln -s /path/to/eden-newgen/study_repo/eden eden
cd ../..
```

## ขั้นที่ 3: ติดตั้ง Python dependencies

```bash
pip install -r applications/eden/requirements.txt
pip install -r applications/eden/optional_requirements.txt

# FastAPI dependencies
pip install -r applications/eden/fastapi_app/requirements.txt
```

## ขั้นที่ 4: ตั้งค่า Eden

```bash
cd applications/eden

# สร้าง config จาก template
cp modules/templates/000_config.py models/000_config.py

# แก้ไข config
sed -i 's/FINISHED_EDITING_CONFIG_FILE = False/FINISHED_EDITING_CONFIG_FILE = True/' models/000_config.py
```

## ขั้นที่ 5: First-time setup (สร้าง database)

```bash
cd /path/to/web2py
python web2py.py -S eden -M -R applications/eden/static/scripts/tools/noop.py
```

รอจนเห็น: `*** FIRST RUN COMPLETE ***`

## ขั้นที่ 6: รันเซิร์ฟเวอร์

### วิธี A: web2py + FastAPI (development)

```bash
# Terminal 1: web2py (legacy)
cd /path/to/web2py
python web2py.py --no_gui -a your_password --port 8000

# Terminal 2: FastAPI (new API)
cd /path/to/eden-newgen/study_repo/eden/fastapi_app
uvicorn main:app --reload --port 8001

# Terminal 3: Next.js (frontend)
cd /path/to/eden-newgen
npm install && npm run dev
```

### วิธี B: Docker (production — แนะนำ)

```bash
cd /path/to/eden-newgen/study_repo/eden
cp .env.example .env
# แก้ .env ใส่ domain + API keys
docker compose -f docker-compose.prod.yml up -d
```

## ขั้นที่ 7: เข้าใช้งาน

```
URL:      http://localhost:8000/eden
Login:    admin@example.com / testing (web2py default)
          admin@example.com / admin123 (FastAPI default)
API Docs: http://localhost:8001/eden/docs
```

---

## โครงสร้างโฟลเดอร์

```
eden-newgen/
├── study_repo/eden/           ← Eden source code (web2py application)
│   ├── controllers/           ← 54 controllers (44 Eden + 10 AI)
│   ├── models/                ← web2py models + zz_ai_tables.py
│   ├── modules/               ← S3 framework + AI services
│   ├── views/                 ← YATL templates
│   ├── fastapi_app/           ← FastAPI backend (49 endpoints)
│   │   ├── app/
│   │   │   ├── api/routers/   ← 12 routers
│   │   │   ├── core/          ← config, database, security
│   │   │   ├── models/        ← SQLAlchemy 2.0 models
│   │   │   ├── schemas/       ← Pydantic v2 schemas
│   │   │   ├── services/      ← AI client, OCR, voice, etc.
│   │   │   └── crud/          ← Data access layer
│   │   ├── tests/             ← 121 tests
│   │   └── requirements.txt
│   ├── Dockerfile
│   ├── docker-compose.prod.yml
│   ├── Caddyfile
│   ├── helm/                  ← Kubernetes chart
│   └── .env.example
├── src/app/                   ← Next.js frontend
│   ├── login/                 ← Login page
│   └── dashboard/             ← Dashboard page
├── package.json               ← Next.js dependencies
├── README.md
└── INSTALL.md                 ← ไฟล์นี้
```

---

## การแก้ปัญหา

### "web2py.py not found"
→ ต้อง clone web2py ก่อน (ขั้นที่ 2)

### "Module not found: structlog"
→ `pip install structlog`

### "Database error"
→ ลบ `databases/storage.db` แล้วรัน first-time setup ใหม่

### "Port already in use"
→ `kill $(lsof -t -i:8000) ; kill $(lsof -t -i:8001)`

---

## ติดต่อ

- GitHub: https://github.com/skpuan07-cmd/eden-newgen
- Issues: https://github.com/skpuan07-cmd/eden-newgen/issues
