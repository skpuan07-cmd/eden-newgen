'use client';
import { useEffect } from 'react';
export default function Home() {
  useEffect(() => {
    const destination = localStorage.getItem('eden_token') ? '/dashboard' : '/login';
    window.location.replace(destination);
  }, []);
  return <div style={{ minHeight: '100vh', background: '#1E3A8A', color: 'white', display: 'flex', alignItems: 'center', justifyContent: 'center' }}>Loading...</div>;
}
