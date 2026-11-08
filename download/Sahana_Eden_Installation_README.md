# Sahana Eden — Installation Summary

วันที่ติดตั้ง: 2026-11-08
สภาพแวดล้อม: Linux sandbox (Python 3.12.16, no Docker)
Source: zip ที่ผู้ใช้อัปโหลด (-sahana-eden-dev.zip)

## โครงสร้างที่ติดตั้ง

```
/home/z/my-project/eden_install/web2py/             ← web2py v3.3.3 (commit 9005306)
├── web2py.py
├── gluon/                                          ← web2py framework
│   └── packages/{pydal, rocket3, yatl}/             ← submodules
└── applications/
    └── eden/                                        ← Sahana Eden (จาก zip ที่อัปโหลด)
        ├── controllers/   (44 ไฟล์)
        ├── models/
        │   └── 000_config.py                       ← แก้ไข FINISHED_EDITING_CONFIG_FILE=True, debug=True
        ├── modules/        (s3 framework + s3db)
        ├── views/
        ├── static/
        ├── databases/
        │   └── storage.db                          ← SQLite, 388 tables, ~3.9 MB
        └── ...
```

## การตั้งค่า

| รายการ | ค่า |
|--------|-----|
| Database | SQLite (`storage.db`) — เหมาะกับ dev/testing |
| Template | `default` (ทุก module เปิดใช้งาน) |
| Debug mode | เปิด (uncompressed CSS/JS) |
| Pre-populate | default base + users + demo data |
| Total tables | 388 |
| Demo users | 2 (admin@example.com / normaluser@example.com — รหัส `testing`) |
| System roles | 9 (ADMIN, AUTHENTICATED, ANONYMOUS, MAP_ADMIN, ORG_ADMIN, ORG_GROUP_ADMIN, ...) |

## วิธีเริ่มต้นระบบ

```bash
# โหมด background (แนะนำ)
/home/z/my-project/scripts/start_eden.sh myPassword daemon

# โหมด foreground (ดู log แบบเรียลไทม์)
/home/z/my-project/scripts/start_eden.sh myPassword
# (กด Ctrl+C เพื่อหยุด)

# หยุดเซิร์ฟเวอร์
/home/z/my-project/scripts/stop_eden.sh

# เปลี่ยน port/IP
EDEN_PORT=8080 EDEN_IP=0.0.0.0 /home/z/my-project/scripts/start_eden.sh myPassword daemon
```

## การเข้าถึง

หลังรัน start script แล้ว เปิดเบราว์เซอร์ไปที่:

- **Eden application**: `http://localhost:8000/eden`
- **web2py admin**: `http://localhost:8000/admin` (ใช้รหัสที่ตั้งตอน start)
- **Login ด้วย demo accounts**:
  - `admin@example.com` / `testing` (system administrator)
  - `normaluser@example.com` / `testing` (regular user)

## ผลการทดสอบ HTTP endpoints (หลัง login)

ทั้งหมดตอบ HTTP 200:

| Endpoint | Title |
|----------|-------|
| /eden/default/index | Sahana Eden |
| /eden/org/organisation | Organizations |
| /eden/org/office | Offices |
| /eden/org/facility | Facilities |
| /eden/pr/person | Persons |
| /eden/hrm/staff | Staff |
| /eden/inv/warehouse | Warehouses |
| /eden/inv/inv_item | Stock in Warehouse |
| /eden/project/project | Projects |
| /eden/asset/asset | Assets |
| /eden/req/req | Requests |
| /eden/cms/post | Posts |

JSON API ตัวอย่าง (ต้อง login ก่อน):
- `GET /eden/pr/person.json` → ส่งคืน `[{"id":1, "first_name":"Admin", ...}]`

## ไฟล์ที่เกี่ยวข้อง

| Path | หน้าที่ |
|------|--------|
| `/home/z/my-project/scripts/start_eden.sh` | สคริปต์เริ่มเซิร์ฟเวอร์ |
| `/home/z/my-project/scripts/stop_eden.sh` | สคริปต์หยุดเซิร์ฟเวอร์ |
| `/home/z/my-project/scripts/test_eden.py` | สคริปต์ทดสอบ endpoints (login + visit หน้าต่าง ๆ + JSON) |
| `/home/z/my-project/eden_install/web2py/applications/eden/models/000_config.py` | ไฟล์ตั้งค่าหลักของ Eden |
| `/home/z/my-project/eden_install/web2py/applications/eden/databases/storage.db` | SQLite database |
| `/tmp/web2py.log` | Log ของเซิร์ฟเวอร์ (ขณะรัน) |

## หมายเหตุเรื่อง Sandbox

- **Sandbox นี้ kill running processes ระหว่างข้อความ** — ทุกครั้งที่ต้องการใช้งาน Eden ต้องรัน `start_eden.sh` ใหม่
- **Database ถูกเก็บไว้ถาวร** ใน `/home/z/my-project/eden_install/...` ข้อมูลไม่หายระหว่าง restart
- ถ้าต้องการ reset database ให้ลบ `databases/storage.db` แล้วรัน first-time setup script ใหม่:
  ```bash
  cd /home/z/my-project/eden_install/web2py
  rm applications/eden/databases/storage.db
  python3 web2py.py -S eden -M -R applications/eden/static/scripts/tools/noop.py
  ```

## ข้อจำกัดที่ควรทราบ

1. **ใช้ SQLite** ไม่ใช่ PostgreSQL/MySQL — เหมาะกับ dev/testing ไม่เหมาะกับ production หลาย user
2. **Optional dependencies ที่ไม่ได้ติดตั้ง** (warning จาก first-run):
   - `GDAL` (Shapefile export)
   - `pyserial` (Serial port modem สำหรับ SMS gateway)
   - `sgmllib3k` (Feed import บน Python 3)
   - `selenium` (Map printing)
3. ถ้าต้องการ production deployment แนะนำติดตั้ง PostgreSQL และใช้ web2py กับ nginx/uWSGI
