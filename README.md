# Eden NewGen

Sahana Eden — Disaster Management Platform with AI extensions

## Features
- FastAPI backend (49 endpoints, 121 tests)
- AI registration: OCR + Voice + Bulk import + NLQ
- JWT auth + 2FA (TOTP)
- PWA offline-first + Push + QR check-in
- Webhooks: Line/SMS/Telegram
- Docker + docker-compose.prod.yml
- Helm chart for Kubernetes
- Next.js frontend: login + dashboard
- 5-phase migration from web2py to FastAPI

## Quick Start
```bash
cp .env.example .env
docker compose -f docker-compose.prod.yml up -d
```

## Login
- Web frontend: http://localhost:3000/login
- Email: admin@example.com
- Password: admin123
- Native Eden login: http://localhost:8000/eden/default/user/login
- http://localhost:8000/ redirects to /eden/default/index

## API
- Swagger UI: /eden/docs (49 endpoints)
- Health: /eden/health/
- OpenAPI: /eden/openapi.json

## AI Modules
1. OCR — Thai ID card auto-fill (OpenThai 2.1)
2. Voice — Thai ASR + intent extraction
3. Bulk Import — Excel/CSV batch processing
4. NLQ — Natural language query in Thai
5. Webhooks — Line/SMS/Telegram bot

## Tech Stack
- FastAPI + SQLAlchemy 2.0 + Pydantic v2
- PostgreSQL 16 + pgvector
- Redis 7 + Celery
- Caddy 2 (auto-HTTPS)
- Next.js 16 + React
- z-ai GLM-4-plus (AI engine)
- Docker + Kubernetes (Helm)

## Roadmap (5 phases, all completed)
1. Foundation — Docker, PostgreSQL, FastAPI, CI/CD
2. AI Registration — OCR, Voice, JWT, duplicate detection
3. Smart Forms — Bulk import, Event Hub, NLQ, Webhooks
4. PWA — Offline-first, Push, QR check-in
5. Production — 2FA, Person CRUD, Helm chart, Security

## Recommendations
- Vector search with pgvector (replace keyword matching)
- Streaming chat (SSE)
- React Native mobile app for field workers
- PostgreSQL read replicas for scale
- WAF + penetration testing
- PDPA DPIA documentation

## License
MIT
