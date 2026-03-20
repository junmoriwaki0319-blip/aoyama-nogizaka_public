export default function middleware(request) {
  const url = new URL(request.url);
  const path = url.pathname;

  // reports.json と known_activists.json への直接アクセスをブロック
  if (path === '/data/reports.json' || path === '/scripts/known_activists.json') {
    return new Response(
      JSON.stringify({ error: '認証が必要です。APIエンドポイントをご利用ください。' }),
      { status: 401, headers: { 'Content-Type': 'application/json' } }
    );
  }
}

export const config = {
  matcher: ['/data/reports.json', '/scripts/known_activists.json'],
};
