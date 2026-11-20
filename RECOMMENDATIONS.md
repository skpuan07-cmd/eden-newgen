# Recommendations for Eden NewGen

## Modernization
1. Docker + docker-compose.prod.yml (done)
2. PostgreSQL 16 + pgvector (done)
3. FastAPI + SQLAlchemy 2.0 (done)
4. Next.js frontend (done)
5. CI/CD with GitHub Actions (done)
6. Helm chart for Kubernetes (done)
7. Tailwind CSS (replace Foundation 5)
8. MapLibre GL (replace OpenLayers)

## AI Enhancement
1. OCR registration with OpenThai 2.1 (done)
2. Voice registration with Thai ASR (done)
3. Bulk import Excel/CSV (done)
4. NLQ natural language query (done)
5. Vector search with pgvector (planned)
6. Streaming chat with SSE (planned)
7. Multi-language support (planned)

## Security
1. JWT auth + 2FA TOTP (done)
2. Rate limiting 4 tiers (done)
3. CSP + HSTS headers (done)
4. Column encryption Fernet (done)
5. Audit logging syslog (done)
6. WAF Cloudflare (planned)
7. Penetration testing (planned)
8. Dependency scanning (done)

## PDPA Compliance
1. Consent enforcement (done)
2. Encrypted sensitive data (done)
3. Image deletion after OCR (done)
4. Audit trail (done)
5. Retention policy automation (planned)
6. Right to be forgotten (planned)
7. DPIA documentation (planned)

## Priority
- Phase 1-5: All completed (121 tests, 49 endpoints)
- Phase 6 (ongoing): Vector search, mobile app, pen test, DPIA
