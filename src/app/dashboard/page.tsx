'use client';
import { useState, useEffect } from 'react';

export default function DashboardPage() {
  const [user, setUser] = useState<any>(null);
  const [persons, setPersons] = useState<any[]>([]);
  const [events, setEvents] = useState<any[]>([]);
  const [health, setHealth] = useState<any>(null);
  const [loaded, setLoaded] = useState(false);

  useEffect(() => {
    const t = localStorage.getItem('eden_token');
    if (!t) { window.location.href = '/login'; return; }
    const u = localStorage.getItem('eden_user');
    if (u) setUser(JSON.parse(u));

    const headers = { Authorization: `Bearer ${t}` };
    Promise.all([
      fetch('/api/eden/health/').then(r => r.json()).then(setHealth).catch(() => {}),
      fetch('/api/eden/api/v2/persons', { headers }).then(r => r.json()).then(setPersons).catch(() => {}),
      fetch('/api/eden/api/v2/event', { headers }).then(r => r.json()).then(setEvents).catch(() => {}),
    ]).finally(() => setLoaded(true));
  }, []);

  if (!loaded) return <div style={{ minHeight: '100vh', display: 'flex', alignItems: 'center', justifyContent: 'center', background: '#1E3A8A', color: 'white' }}>กำลังโหลด...</div>;

  return (
    <div style={{ minHeight: '100vh', background: '#F8FAFC', fontFamily: 'system-ui, sans-serif' }}>
      <div style={{ background: '#1E3A8A', color: 'white', padding: '1rem 2rem', display: 'flex', justifyContent: 'space-between', alignItems: 'center' }}>
        <div style={{ display: 'flex', alignItems: 'center', gap: '0.5rem' }}>
          <span style={{ fontSize: '1.5rem' }}>🌻</span>
          <span style={{ fontWeight: 600 }}>Sahana Eden</span>
        </div>
        <div style={{ display: 'flex', alignItems: 'center', gap: '1rem' }}>
          <span style={{ fontSize: '0.9rem', opacity: 0.8 }}>{user?.email}</span>
          <button onClick={() => { localStorage.clear(); window.location.href = '/login'; }} style={{ background: 'rgba(255,255,255,0.2)', border: 'none', color: 'white', padding: '0.4rem 1rem', borderRadius: '6px', cursor: 'pointer' }}>ออกจากระบบ</button>
        </div>
      </div>
      <div style={{ maxWidth: '1200px', margin: '0 auto', padding: '2rem' }}>
        <div style={{ display: 'grid', gridTemplateColumns: 'repeat(auto-fill, minmax(220px, 1fr))', gap: '1rem', marginBottom: '2rem' }}>
          <div style={{ background: 'white', borderRadius: '12px', padding: '1.5rem', boxShadow: '0 1px 3px rgba(0,0,0,0.08)' }}>
            <div style={{ fontSize: '0.85rem', color: '#64748B' }}>สถานะระบบ</div>
            <div style={{ fontSize: '1.5rem', fontWeight: 700, color: health?.status === 'ok' ? '#16A34A' : '#D97706' }}>{health?.status || '...'}</div>
          </div>
          <div style={{ background: 'white', borderRadius: '12px', padding: '1.5rem', boxShadow: '0 1px 3px rgba(0,0,0,0.08)' }}>
            <div style={{ fontSize: '0.85rem', color: '#64748B' }}>บุคคลในระบบ</div>
            <div style={{ fontSize: '1.5rem', fontWeight: 700, color: '#1E3A8A' }}>{persons.length} คน</div>
          </div>
          <div style={{ background: 'white', borderRadius: '12px', padding: '1.5rem', boxShadow: '0 1px 3px rgba(0,0,0,0.08)' }}>
            <div style={{ fontSize: '0.85rem', color: '#64748B' }}>เหตุการณ์ภัยพิบัติ</div>
            <div style={{ fontSize: '1.5rem', fontWeight: 700, color: '#DC2626' }}>{events.length} เหตุการณ์</div>
          </div>
        </div>
        <div style={{ background: 'white', borderRadius: '12px', padding: '1.5rem', marginBottom: '2rem', boxShadow: '0 1px 3px rgba(0,0,0,0.08)' }}>
          <h2 style={{ fontSize: '1.1rem', marginBottom: '1rem' }}>เมนูหลัก</h2>
          <div style={{ display: 'grid', gridTemplateColumns: 'repeat(auto-fill, minmax(200px, 1fr))', gap: '0.75rem' }}>
            <a href="/api/eden/docs" style={{ padding: '1rem', background: '#F8FAFC', borderRadius: '8px', textDecoration: 'none', color: '#334155', border: '1px solid #E2E8F0' }}>📖 API Docs</a>
            <a href="/api/eden/openapi.json" style={{ padding: '1rem', background: '#F8FAFC', borderRadius: '8px', textDecoration: 'none', color: '#334155', border: '1px solid #E2E8F0' }}>📄 OpenAPI Spec</a>
            <a href="/api/eden/health/" style={{ padding: '1rem', background: '#F8FAFC', borderRadius: '8px', textDecoration: 'none', color: '#334155', border: '1px solid #E2E8F0' }}>❤️ Health Check</a>
            <a href="/api/eden/manifest.json" style={{ padding: '1rem', background: '#F8FAFC', borderRadius: '8px', textDecoration: 'none', color: '#334155', border: '1px solid #E2E8F0' }}>📱 PWA</a>
          </div>
        </div>
        <div style={{ background: 'white', borderRadius: '12px', padding: '1.5rem', boxShadow: '0 1px 3px rgba(0,0,0,0.08)' }}>
          <h2 style={{ fontSize: '1.1rem', marginBottom: '1rem' }}>บุคคลในระบบ</h2>
          {persons.length === 0 ? <p style={{ color: '#94A3B8' }}>ยังไม่มีข้อมูล — ลงทะเบียนได้ที่ POST /eden/api/v2/register/staff</p> : (
            <table style={{ width: '100%', borderCollapse: 'collapse' }}>
              <thead><tr style={{ borderBottom: '2px solid #E2E8F0' }}>
                <th style={{ textAlign: 'left', padding: '0.5rem', fontSize: '0.85rem', color: '#64748B' }}>ID</th>
                <th style={{ textAlign: 'left', padding: '0.5rem', fontSize: '0.85rem', color: '#64748B' }}>ชื่อ</th>
                <th style={{ textAlign: 'left', padding: '0.5rem', fontSize: '0.85rem', color: '#64748B' }}>นามสกุล</th>
                <th style={{ textAlign: 'left', padding: '0.5rem', fontSize: '0.85rem', color: '#64748B' }}>เพศ</th>
                <th style={{ textAlign: 'left', padding: '0.5rem', fontSize: '0.85rem', color: '#64748B' }}>บัตร ปชช.</th>
              </tr></thead>
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
