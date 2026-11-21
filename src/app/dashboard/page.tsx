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
        <button onClick={() => { localStorage.clear(); window.location.href = '/login'; }} style={{ background: 'rgba(255,255,255,0.2)', border: 'none', color: 'white', padding: '0.4rem 1rem', borderRadius: '6px', cursor: 'pointer' }}>ออกจากระบบ</button>
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
            <div style={{ fontSize: '0.85rem', color: '#64748B' }}>เหตุการณ์</div>
            <div style={{ fontSize: '1.5rem', fontWeight: 700, color: '#DC2626' }}>{events.length}</div>
          </div>
        </div>
        <div style={{ background: 'white', borderRadius: '12px', padding: '1.5rem', boxShadow: '0 1px 3px rgba(0,0,0,0.08)' }}>
          <h2 style={{ fontSize: '1.1rem', marginBottom: '1rem' }}>บุคคลในระบบ</h2>
          {persons.length === 0 ? <p style={{ color: '#94A3B8' }}>ยังไม่มีข้อมูล</p> : (
            <table style={{ width: '100%', borderCollapse: 'collapse' }}>
              <thead><tr style={{ borderBottom: '2px solid #E2E8F0' }}>
                <th style={{ textAlign: 'left', padding: '0.5rem', fontSize: '0.85rem', color: '#64748B' }}>ID</th>
                <th style={{ textAlign: 'left', padding: '0.5rem', fontSize: '0.85rem', color: '#64748B' }}>ชื่อ</th>
                <th style={{ textAlign: 'left', padding: '0.5rem', fontSize: '0.85rem', color: '#64748B' }}>นามสกุล</th>
              </tr></thead>
              <tbody>{persons.map(p => (<tr key={p.id} style={{ borderBottom: '1px solid #F1F5F9' }}>
                <td style={{ padding: '0.5rem' }}>{p.id}</td>
                <td style={{ padding: '0.5rem' }}>{p.first_name}</td>
                <td style={{ padding: '0.5rem' }}>{p.last_name}</td>
              </tr>))}</tbody>
            </table>
          )}
        </div>
      </div>
    </div>
  );
}
