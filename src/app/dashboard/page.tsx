'use client';

import { useState, useEffect } from 'react';

export default function DashboardPage() {
  const [user, setUser] = useState<any>(null);
  const [token, setToken] = useState('');
  const [health, setHealth] = useState<any>(null);
  const [persons, setPersons] = useState<any[]>([]);
  const [events, setEvents] = useState<any[]>([]);

  useEffect(() => {
    const t = localStorage.getItem('eden_token');
    const u = localStorage.getItem('eden_user');
    if (!t) {
      window.location.href = '/login';
      return;
    }
    setToken(t);
    if (u) setUser(JSON.parse(u));

    // Fetch data
    const headers = { Authorization: `Bearer ${t}` };

    fetch('/eden/health/').then(r => r.json()).then(setHealth).catch(() => {});
    fetch('/eden/api/v2/persons', { headers }).then(r => r.json()).then(setPersons).catch(() => {});
    fetch('/eden/api/v2/event', { headers }).then(r => r.json()).then(setEvents).catch(() => {});
  }, []);

  const logout = () => {
    localStorage.removeItem('eden_token');
    localStorage.removeItem('eden_user');
    window.location.href = '/login';
  };

  if (!token) return null;

  return (
    <div style={{ minHeight: '100vh', background: '#F8FAFC', fontFamily: 'system-ui, sans-serif' }}>
      {/* Header */}
      <div style={{
        background: '#1E3A8A', color: 'white', padding: '1rem 2rem',
        display: 'flex', justifyContent: 'space-between', alignItems: 'center',
      }}>
        <div style={{ display: 'flex', alignItems: 'center', gap: '0.5rem' }}>
          <span style={{ fontSize: '1.5rem' }}>🌻</span>
          <span style={{ fontSize: '1.2rem', fontWeight: 600 }}>Sahana Eden</span>
        </div>
        <div style={{ display: 'flex', alignItems: 'center', gap: '1rem' }}>
          <span style={{ fontSize: '0.9rem', opacity: 0.8 }}>
            {user?.email || 'admin'}
          </span>
          <button onClick={logout} style={{
            background: 'rgba(255,255,255,0.2)', border: 'none', color: 'white',
            padding: '0.4rem 1rem', borderRadius: '6px', cursor: 'pointer', fontSize: '0.85rem',
          }}>
            ออกจากระบบ
          </button>
        </div>
      </div>

      {/* Content */}
      <div style={{ maxWidth: '1200px', margin: '0 auto', padding: '2rem' }}>
        {/* Stats cards */}
        <div style={{ display: 'grid', gridTemplateColumns: 'repeat(auto-fill, minmax(250px, 1fr))', gap: '1rem', marginBottom: '2rem' }}>
          {/* Health */}
          <div style={{ background: 'white', borderRadius: '12px', padding: '1.5rem', boxShadow: '0 1px 3px rgba(0,0,0,0.08)' }}>
            <div style={{ fontSize: '0.85rem', color: '#64748B', marginBottom: '0.5rem' }}>สถานะระบบ</div>
            <div style={{ fontSize: '1.5rem', fontWeight: 700, color: health?.status === 'ok' ? '#16A34A' : '#D97706' }}>
              {health?.status || '...'}
            </div>
            <div style={{ fontSize: '0.75rem', color: '#94A3B8', marginTop: '0.25rem' }}>
              v{health?.version || '...'}
            </div>
          </div>

          {/* Persons */}
          <div style={{ background: 'white', borderRadius: '12px', padding: '1.5rem', boxShadow: '0 1px 3px rgba(0,0,0,0.08)' }}>
            <div style={{ fontSize: '0.85rem', color: '#64748B', marginBottom: '0.5rem' }}>บุคคลในระบบ</div>
            <div style={{ fontSize: '1.5rem', fontWeight: 700, color: '#1E3A8A' }}>
              {persons.length}
            </div>
            <div style={{ fontSize: '0.75rem', color: '#94A3B8', marginTop: '0.25rem' }}>คน</div>
          </div>

          {/* Events */}
          <div style={{ background: 'white', borderRadius: '12px', padding: '1.5rem', boxShadow: '0 1px 3px rgba(0,0,0,0.08)' }}>
            <div style={{ fontSize: '0.85rem', color: '#64748B', marginBottom: '0.5rem' }}>เหตุการณ์ภัยพิบัติ</div>
            <div style={{ fontSize: '1.5rem', fontWeight: 700, color: '#DC2626' }}>
              {events.length}
            </div>
            <div style={{ fontSize: '0.75rem', color: '#94A3B8', marginTop: '0.25rem' }}>เหตุการณ์</div>
          </div>

          {/* Endpoints */}
          <div style={{ background: 'white', borderRadius: '12px', padding: '1.5rem', boxShadow: '0 1px 3px rgba(0,0,0,0.08)' }}>
            <div style={{ fontSize: '0.85rem', color: '#64748B', marginBottom: '0.5rem' }}>API Endpoints</div>
            <div style={{ fontSize: '1.5rem', fontWeight: 700, color: '#7C3AED' }}>49</div>
            <div style={{ fontSize: '0.75rem', color: '#94A3B8', marginTop: '0.25rem' }}>routes</div>
          </div>
        </div>

        {/* Quick actions */}
        <div style={{ background: 'white', borderRadius: '12px', padding: '1.5rem', marginBottom: '2rem', boxShadow: '0 1px 3px rgba(0,0,0,0.08)' }}>
          <h2 style={{ fontSize: '1.1rem', color: '#0F172A', marginBottom: '1rem' }}>เมนูหลัก</h2>
          <div style={{ display: 'grid', gridTemplateColumns: 'repeat(auto-fill, minmax(200px, 1fr))', gap: '0.75rem' }}>
            {[
              { label: '📖 API Docs', href: '/eden/docs', desc: 'Swagger UI' },
              { label: '👥 ลงทะเบียนเจ้าหน้าที่', href: '/eden/api/v2/register/staff', desc: 'Register staff' },
              { label: '🤳 สแกนบัตรประชาชน', href: '/eden/api/v2/register/ocr', desc: 'OCR registration' },
              { label: '🎤 ลงทะเบียนด้วยเสียง', href: '/eden/api/v2/register/voice', desc: 'Voice registration' },
              { label: '📊 นำเข้าข้อมูลกลุ่ม', href: '/eden/api/v2/register/bulk', desc: 'Bulk import' },
              { label: '❓ ถามคำถาม (AI)', href: '/eden/api/v2/ai_query/ask', desc: 'NLQ' },
              { label: '🚨 สร้างเหตุการณ์ภัยพิบัติ', href: '/eden/api/v2/event', desc: 'Create event' },
              { label: '📱 PWA Manifest', href: '/eden/manifest.json', desc: 'Install app' },
            ].map(item => (
              <a key={item.href} href={item.href} style={{
                display: 'block', padding: '1rem', background: '#F8FAFC',
                borderRadius: '8px', textDecoration: 'none', color: '#334155',
                border: '1px solid #E2E8F0', transition: 'all 0.15s',
              }}>
                <div style={{ fontWeight: 600, fontSize: '0.9rem' }}>{item.label}</div>
                <div style={{ fontSize: '0.75rem', color: '#94A3B8', marginTop: '0.25rem' }}>{item.desc}</div>
              </a>
            ))}
          </div>
        </div>

        {/* Persons list */}
        <div style={{ background: 'white', borderRadius: '12px', padding: '1.5rem', boxShadow: '0 1px 3px rgba(0,0,0,0.08)' }}>
          <h2 style={{ fontSize: '1.1rem', color: '#0F172A', marginBottom: '1rem' }}>บุคคลในระบบ</h2>
          {persons.length === 0 ? (
            <p style={{ color: '#94A3B8' }}>ยังไม่มีข้อมูล</p>
          ) : (
            <table style={{ width: '100%', borderCollapse: 'collapse' }}>
              <thead>
                <tr style={{ borderBottom: '2px solid #E2E8F0' }}>
                  <th style={{ textAlign: 'left', padding: '0.5rem', fontSize: '0.85rem', color: '#64748B' }}>ID</th>
                  <th style={{ textAlign: 'left', padding: '0.5rem', fontSize: '0.85rem', color: '#64748B' }}>ชื่อ</th>
                  <th style={{ textAlign: 'left', padding: '0.5rem', fontSize: '0.85rem', color: '#64748B' }}>นามสกุล</th>
                  <th style={{ textAlign: 'left', padding: '0.5rem', fontSize: '0.85rem', color: '#64748B' }}>เพศ</th>
                  <th style={{ textAlign: 'left', padding: '0.5rem', fontSize: '0.85rem', color: '#64748B' }}>บัตรประชาชน</th>
                </tr>
              </thead>
              <tbody>
                {persons.map(p => (
                  <tr key={p.id} style={{ borderBottom: '1px solid #F1F5F9' }}>
                    <td style={{ padding: '0.5rem', fontSize: '0.9rem' }}>{p.id}</td>
                    <td style={{ padding: '0.5rem', fontSize: '0.9rem' }}>{p.first_name}</td>
                    <td style={{ padding: '0.5rem', fontSize: '0.9rem' }}>{p.last_name}</td>
                    <td style={{ padding: '0.5rem', fontSize: '0.9rem' }}>{p.gender || '-'}</td>
                    <td style={{ padding: '0.5rem', fontSize: '0.9rem' }}>{p.has_national_id ? '✓' : '-'}</td>
                  </tr>
                ))}
              </tbody>
            </table>
          )}
        </div>
      </div>
    </div>
  );
}
