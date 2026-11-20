'use client';

import { useEffect } from 'react';

export default function Home() {
  useEffect(() => {
    const token = localStorage.getItem('eden_token');
    if (token) {
      window.location.href = '/dashboard';
    } else {
      window.location.href = '/login';
    }
  }, []);

  return (
    <div style={{
      minHeight: '100vh', display: 'flex', alignItems: 'center', justifyContent: 'center',
      background: '#1E3A8A', color: 'white', fontFamily: 'system-ui, sans-serif',
    }}>
      <p>Loading...</p>
    </div>
  );
}
