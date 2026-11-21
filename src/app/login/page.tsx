'use client';
import { useState } from 'react';

export default function LoginPage() {
  const [email, setEmail] = useState('admin@example.com');
  const [password, setPassword] = useState('admin123');
  const [loading, setLoading] = useState(false);
  const [error, setError] = useState('');
  const [success, setSuccess] = useState(false);

  const handleLogin = async (e: React.FormEvent) => {
    e.preventDefault();
    setLoading(true);
    setError('');
    setSuccess(false);

    try {
      // Use RELATIVE URL — goes through Next.js rewrite proxy to FastAPI
      const params = new URLSearchParams();
      params.append('username', email);
      params.append('password', password);

      const res = await fetch('/api/eden/api/v2/auth/login', {
        method: 'POST',
        headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
        body: params.toString(),
      });

      const data = await res.json();

      if (res.ok && data.access_token) {
        localStorage.setItem('eden_token', data.access_token);
        localStorage.setItem('eden_user', JSON.stringify(data.user));
        setSuccess(true);
        setTimeout(() => { window.location.href = '/dashboard'; }, 800);
      } else {
        setError(data.detail || 'เข้าสู่ระบบไม่สำเร็จ');
      }
    } catch (err) {
      setError('ไม่สามารถเชื่อมต่อเซิร์ฟเวอร์ — พิมพ์ "run eden" เพื่อรีสตาร์ท');
    } finally {
      setLoading(false);
    }
  };

  return (
    <div style={{ minHeight: '100vh', display: 'flex', alignItems: 'center', justifyContent: 'center', background: 'linear-gradient(135deg, #1E3A8A, #0E7490)', fontFamily: 'system-ui, sans-serif' }}>
      <div style={{ background: 'white', borderRadius: '16px', padding: '2.5rem', width: '100%', maxWidth: '400px', boxShadow: '0 20px 60px rgba(0,0,0,0.3)' }}>
        <div style={{ textAlign: 'center', marginBottom: '2rem' }}>
          <div style={{ fontSize: '3rem' }}>🌻</div>
          <h1 style={{ color: '#1E3A8A', margin: '0.5rem 0 0.25rem' }}>Sahana Eden</h1>
          <p style={{ color: '#64748B', fontSize: '0.9rem', margin: 0 }}>ระบบจัดการภัยพิบัติ</p>
        </div>
        <form onSubmit={handleLogin}>
          <div style={{ marginBottom: '1rem' }}>
            <label style={{ display: 'block', fontSize: '0.85rem', fontWeight: 600, color: '#64748B', marginBottom: '0.25rem' }}>อีเมล</label>
            <input type="email" value={email} onChange={e => setEmail(e.target.value)} style={{ width: '100%', padding: '0.75rem', border: '1px solid #E2E8F0', borderRadius: '8px', fontSize: '1rem', boxSizing: 'border-box' }} required />
          </div>
          <div style={{ marginBottom: '1.5rem' }}>
            <label style={{ display: 'block', fontSize: '0.85rem', fontWeight: 600, color: '#64748B', marginBottom: '0.25rem' }}>รหัสผ่าน</label>
            <input type="password" value={password} onChange={e => setPassword(e.target.value)} style={{ width: '100%', padding: '0.75rem', border: '1px solid #E2E8F0', borderRadius: '8px', fontSize: '1rem', boxSizing: 'border-box' }} required />
          </div>
          {error && <div style={{ background: '#FEE2E2', color: '#991B1B', padding: '0.75rem', borderRadius: '8px', fontSize: '0.85rem', marginBottom: '1rem' }}>⚠ {error}</div>}
          {success && <div style={{ background: '#D1FAE5', color: '#065F46', padding: '0.75rem', borderRadius: '8px', fontSize: '0.85rem', marginBottom: '1rem' }}>✅ เข้าสู่ระบบสำเร็จ! กำลังไปหน้า Dashboard...</div>}
          <button type="submit" disabled={loading} style={{ width: '100%', padding: '0.75rem', background: '#1E3A8A', color: 'white', border: 'none', borderRadius: '8px', fontSize: '1rem', fontWeight: 600, cursor: loading ? 'wait' : 'pointer', opacity: loading ? 0.7 : 1 }}>
            {loading ? 'กำลังเข้าสู่ระบบ...' : 'เข้าสู่ระบบ'}
          </button>
        </form>
        <div style={{ marginTop: '1.5rem', padding: '0.75rem', background: '#F8FAFC', borderRadius: '8px', fontSize: '0.8rem', color: '#64748B', textAlign: 'center' }}>
          Demo: admin@example.com / admin123
        </div>
      </div>
    </div>
  );
}
