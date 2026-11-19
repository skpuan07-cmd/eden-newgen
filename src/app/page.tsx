export default function Home() {
  return (
    <div style={{ 
      display: 'flex', flexDirection: 'column', alignItems: 'center',
      justifyContent: 'center', minHeight: '100vh',
      fontFamily: 'system-ui, sans-serif', background: '#1E3A8A',
      color: 'white', gap: '2rem'
    }}>
      <h1 style={{ fontSize: '2.5rem', margin: 0 }}>🌻 Sahana Eden</h1>
      <p style={{ opacity: 0.8, fontSize: '1.2rem' }}>Disaster Management Platform with AI</p>
      <div style={{ display: 'flex', gap: '1rem', flexWrap: 'wrap', justifyContent: 'center' }}>
        <a href="/eden/docs" style={{ padding: '0.75rem 1.5rem', background: 'white', color: '#1E3A8A', borderRadius: '8px', textDecoration: 'none', fontWeight: 600 }}>📖 API Docs (Swagger)</a>
        <a href="/eden/health/" style={{ padding: '0.75rem 1.5rem', background: 'transparent', border: '2px solid white', color: 'white', borderRadius: '8px', textDecoration: 'none', fontWeight: 600 }}>❤️ Health Check</a>
      </div>
      <p style={{ opacity: 0.6, fontSize: '0.9rem', marginTop: '2rem' }}>49 API endpoints · 121 tests · Phase 5</p>
    </div>
  );
}
