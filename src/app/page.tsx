'use client';
import { useEffect } from 'react';
export default function Home() {
  useEffect(() => {
    window.location.href = localStorage.getItem('eden_token') ? '/dashboard' : '/login';
  }, []);
  return <div style={{ minHeight: '100vh', background: '#1E3A8A', color: 'white', display: 'flex', alignItems: 'center', justifyContent: 'center' }}>Loading...</div>;
}
