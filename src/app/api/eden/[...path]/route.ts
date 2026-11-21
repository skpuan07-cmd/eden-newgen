import { NextRequest, NextResponse } from 'next/server';

const BACKEND = 'http://127.0.0.1:8001';

<<<<<<< HEAD
async function proxy(req: NextRequest, { params }: { params: { path: string[] } }) {
  const path = params.path.join('/');
=======
async function proxy(req: NextRequest, { params }: { params: Promise<{ path: string[] }> }) {
  const { path: segments } = await params;
  const path = segments.join('/');
>>>>>>> 4d5abf907c02ad7517171d310ee1e0904ed49c7e
  const search = req.nextUrl.search;
  const url = `${BACKEND}/eden/${path}${search}`;
  
  // Build headers — preserve content-type
  const h = new Headers();
  const ct = req.headers.get('content-type');
  if (ct) h.set('content-type', ct);
  const auth = req.headers.get('authorization');
  if (auth) h.set('authorization', auth);
  
  const init: RequestInit = { method: req.method, headers: h };
  
  if (req.method !== 'GET' && req.method !== 'HEAD') {
    init.body = await req.text();
  }
  
  try {
    const r = await fetch(url, init);
    const body = await r.text();
    const rh = new Headers();
    r.headers.forEach((v, k) => { 
      if (k !== 'transfer-encoding' && k !== 'content-encoding') rh.set(k, v); 
    });
    return new NextResponse(body, { status: r.status, headers: rh });
<<<<<<< HEAD
  } catch (e: any) {
    return NextResponse.json({ error: 'Backend unavailable', detail: e.message }, { status: 502 });
=======
  } catch (error) {
    return NextResponse.json(
      {
        error: 'Backend unavailable',
        detail: error instanceof Error ? error.message : 'Unknown backend error',
      },
      { status: 502 },
    );
>>>>>>> 4d5abf907c02ad7517171d310ee1e0904ed49c7e
  }
}

export const GET = proxy;
export const POST = proxy;
export const PATCH = proxy;
export const DELETE = proxy;
