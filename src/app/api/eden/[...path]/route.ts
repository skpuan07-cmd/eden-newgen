import { NextRequest, NextResponse } from 'next/server';

const BACKEND_URL = 'http://127.0.0.1:8001';

export async function GET(request: NextRequest, { params }: { params: { path: string[] } }) {
  return proxy(request, params);
}
export async function POST(request: NextRequest, { params }: { params: { path: string[] } }) {
  return proxy(request, params);
}
export async function PATCH(request: NextRequest, { params }: { params: { path: string[] } }) {
  return proxy(request, params);
}
export async function DELETE(request: NextRequest, { params }: { params: { path: string[] } }) {
  return proxy(request, params);
}

async function proxy(request: NextRequest, params: { path: string[] }) {
  const path = params.path.join('/');
  const url = `${BACKEND_URL}/eden/${path}${request.nextUrl.search}`;
  
  const headers = new Headers(request.headers);
  headers.delete('host');
  
  const init: RequestInit = {
    method: request.method,
    headers,
  };
  
  if (request.method !== 'GET' && request.method !== 'HEAD') {
    const contentType = request.headers.get('content-type') || '';
    if (contentType.includes('application/json')) {
      init.body = await request.text();
    } else if (contentType.includes('multipart/form-data')) {
      init.body = await request.blob();
    } else {
      init.body = await request.text();
    }
  }
  
  try {
    const res = await fetch(url, init);
    const data = await res.text();
    const responseHeaders = new Headers();
    res.headers.forEach((v, k) => responseHeaders.set(k, v));
    return new NextResponse(data, { status: res.status, headers: responseHeaders });
  } catch (e) {
    return NextResponse.json({ error: 'Backend unavailable' }, { status: 502 });
  }
}
