'use client';

import { useState } from 'react';

export default function LoginPage() {
  const [email, setEmail] = useState('admin@eden.local');
  const [password, setPassword] = useState('admin123');
  const [loading, setLoading] = useState(false);
  const [error, setError] = useState('');
  const [token, setToken] = useState('');

  const handleLogin = async (e: React.FormEvent) => {
    e.preventDefault();
    setLoading(true);
    setError('');

    try {
      const formData = new FormData();
      formData.append('username', email);
      formData.append('password', password);

      const res = await fetch('/eden/api/v2/auth/login', {
        method: 'POST',
        body: formData,
      });

      if (res.ok) {
        const data = await res.json();
        setToken(data.access_token);
        localStorage.setItem('eden_token', data.access_token);
        localStorage.setItem('eden_user', JSON.stringify(data.user));
        window.location.href = '/dashboard';
      } else {
        const err = await res.json();
        setError(err.detail || 'Login failed');
      }
    } catch (err) {
      setError('Network error');
    } finally {
      setLoading(false);
    }
  };

  return (
    <div style={{
      minHeight: '100vh', display: 'flex', alignItems: 'center', justifyContent: 'center',
      background: 'linear-gradient(135deg, #1E3A8A 0%, #0E7490 100%)',
      fontFamily: 'system-ui, -apple-system, sans-serif',
    }}>
      <div style={{
        background: 'white', borderRadius: '16px', padding: '2.5rem',
        width: '100%', maxWidth: '400px', boxShadow: '0 20px 60px rgba(0,0,0,0.3)',
      }}>
        <div style={{ textAlign: 'center', marginBottom: '2rem' }}>
          <div style={{ fontSize: '3rem', marginBottom: '0.5rem' }}>🌻</div>
          <h1 style={{ fontSize: '1.5rem', color: '#1E3A8A', margin: '0 0 0.25rem' }}>
            Sahana Eden
          </h1>
          <p style={{ color: '#64748B', fontSize: '0.9rem', margin: 0 }}>
            ระบบจัดการภัยพิบัติ
          </p>
        </div>

        <form onSubmit={handleLogin}>
          <div style={{ marginBottom: '1rem' }}>
            <label style={{ display: 'block', fontSize: '0.85rem', fontWeight: 600, color: '#64748B', marginBottom: '0.25rem' }}>
              อีเมล
            </label>
            <input
              type="email"
              value={email}
              onChange={(e) => setEmail(e.target.value)}
              style={{
                width: '100%', padding: '0.75rem', border: '1px solid #E2E8F0',
                borderRadius: '8px', fontSize: '1rem', boxSizing: 'border-box',
              }}
              required
            />
          </div>

          <div style={{ marginBottom: '1.5rem' }}>
            <label style={{ display: 'block', fontSize: '0.85rem', fontWeight: 600, color: '#64748B', marginBottom: '0.25rem' }}>
              รหัสผ่าน
            </label>
            <input
              type="password"
              value={password}
              onChange={(e) => setPassword(e.target.value)}
              style={{
                width: '100%', padding: '0.75rem', border: '1px solid #E2E8F0',
                borderRadius: '8px', fontSize: '1rem', boxSizing: 'border-box',
              }}
              required
            />
          </div>

          {error && (
            <div style={{
              background: '#FEE2E2', color: '#991B1B', padding: '0.75rem',
              borderRadius: '8px', fontSize: '0.85rem', marginBottom: '1rem',
            }}>
              {error}
            </div>
          )}

          <button
            type="submit"
            disabled={loading}
            style={{
              width: '100%', padding: '0.75rem', background: '#1E3A8A',
              color: 'white', border: 'none', borderRadius: '8px',
              fontSize: '1rem', fontWeight: 600, cursor: loading ? 'not-allowed' : 'pointer',
              opacity: loading ? 0.7 : 1,
            }}
          >
            {loading ? 'กำลังเข้าสู่ระบบ...' : 'เข้าสู่ระบบ'}
          </button>
        </form>

        <div style={{
          marginTop: '1.5rem', padding: '0.75rem', background: '#F8FAFC',
          borderRadius: '8px', fontSize: '0.8rem', color: '#64748B', textAlign: 'center',
        }}>
          Demo: admin@eden.local / admin123
        </div>

        <div style={{ marginTop: '1rem', textAlign: 'center' }}>
          <a href="/eden/docs" style={{ color: '#0E7490', fontSize: '0.85rem', textDecoration: 'none' }}>
            📖 API Docs (Swagger)
          </a>
        </div>
      </div>
    </div>
  );
}
