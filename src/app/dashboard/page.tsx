'use client';
import { useState, useEffect } from 'react';

type Person = {
  id: number;
  first_name: string;
  last_name: string;
};

type EdenEvent = {
  id: number;
  name: string;
};

type Health = {
  status: string;
};

function isRecord(value: unknown): value is Record<string, unknown> {
  return typeof value === 'object' && value !== null;
}

function isPerson(value: unknown): value is Person {
  return isRecord(value) &&
    typeof value.id === 'number' &&
    typeof value.first_name === 'string' &&
    typeof value.last_name === 'string';
}

function isEvent(value: unknown): value is EdenEvent {
  return isRecord(value) &&
    typeof value.id === 'number' &&
    typeof value.name === 'string';
}

function isHealth(value: unknown): value is Health {
  return isRecord(value) && typeof value.status === 'string';
}

class UnauthorizedError extends Error {}

async function fetchEden<T>(
  path: string,
  isExpected: (value: unknown) => value is T,
  token?: string,
): Promise<T> {
  const headers = token ? { Authorization: `Bearer ${token}` } : undefined;
  const response = await fetch(path, { headers });

  if (response.status === 401) {
    throw new UnauthorizedError('เซสชันหมดอายุ กรุณาเข้าสู่ระบบอีกครั้ง');
  }

  let data: unknown;
  try {
    data = await response.json();
  } catch {
    throw new Error('เซิร์ฟเวอร์ตอบกลับข้อมูลไม่ถูกต้อง');
  }

  if (!response.ok) {
    const detail = isRecord(data) && typeof data.detail === 'string'
      ? data.detail
      : `โหลดข้อมูลไม่สำเร็จ (${response.status})`;
    throw new Error(detail);
  }

  if (!isExpected(data)) {
    throw new Error('รูปแบบข้อมูลจากเซิร์ฟเวอร์ไม่ถูกต้อง');
  }

  return data;
}

export default function DashboardPage() {
  const [persons, setPersons] = useState<Person[]>([]);
  const [events, setEvents] = useState<EdenEvent[]>([]);
  const [health, setHealth] = useState<Health | null>(null);
  const [loaded, setLoaded] = useState(false);
  const [loadError, setLoadError] = useState('');

  useEffect(() => {
    let cancelled = false;
    const token = localStorage.getItem('eden_token');

    if (!token) {
      window.location.replace('/login');
      return () => { cancelled = true; };
    }

    const loadDashboard = async () => {
      try {
        const [healthData, personData, eventData] = await Promise.all([
          fetchEden('/api/eden/health/', isHealth),
          fetchEden(
            '/api/eden/api/v2/persons',
            (value): value is Person[] => Array.isArray(value) && value.every(isPerson),
            token,
          ),
          fetchEden(
            '/api/eden/api/v2/event',
            (value): value is EdenEvent[] => Array.isArray(value) && value.every(isEvent),
            token,
          ),
        ]);

        if (!cancelled) {
          setHealth(healthData);
          setPersons(personData);
          setEvents(eventData);
        }
      } catch (error) {
        if (error instanceof UnauthorizedError) {
          localStorage.removeItem('eden_token');
          localStorage.removeItem('eden_user');
          window.location.replace('/login');
          return;
        }
        if (!cancelled) {
          setLoadError(error instanceof Error ? error.message : 'โหลดข้อมูลไม่สำเร็จ');
        }
      } finally {
        if (!cancelled) setLoaded(true);
      }
    };

    void loadDashboard();
    return () => { cancelled = true; };
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
          <button onClick={() => {
            localStorage.removeItem('eden_token');
            localStorage.removeItem('eden_user');
            window.location.replace('/login');
          }} style={{ background: 'rgba(255,255,255,0.2)', border: 'none', color: 'white', padding: '0.4rem 1rem', borderRadius: '6px', cursor: 'pointer' }}>ออกจากระบบ</button>
        </div>
      </div>
      <div style={{ maxWidth: '1200px', margin: '0 auto', padding: '2rem' }}>
        {loadError && <div role="alert" style={{ background: '#FEE2E2', color: '#991B1B', padding: '0.75rem', borderRadius: '8px', marginBottom: '1rem' }}>{loadError}</div>}
        <div style={{ display: 'grid', gridTemplateColumns: 'repeat(auto-fill, minmax(220px, 1fr))', gap: '1rem', marginBottom: '2rem' }}>
          <div style={{ background: 'white', borderRadius: '12px', padding: '1.5rem', boxShadow: '0 1px 3px rgba(0,0,0,0.08)' }}>
            <div style={{ fontSize: '0.85rem', color: '#64748B' }}>สถานะระบบ</div>
            <div style={{ fontSize: '1.5rem', fontWeight: 700, color: health?.status === 'ok' ? '#16A34A' : '#D97706' }}>{health?.status || 'ไม่ทราบ'}</div>
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
