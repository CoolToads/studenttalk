const allowedOrigin = 'https://jonathanp123yay.github.io'

function json(data, status = 200, origin = allowedOrigin) {
  return new Response(JSON.stringify(data), {
    status,
    headers: {
      'content-type': 'application/json; charset=utf-8',
      'access-control-allow-origin': origin,
      'access-control-allow-headers': 'content-type, x-user-email, x-user-name',
      'access-control-allow-methods': 'GET, POST, OPTIONS',
      'vary': 'Origin',
    },
  })
}

function getOrigin(request) {
  const origin = request.headers.get('Origin')
  return origin === allowedOrigin || origin?.startsWith('http://localhost:') ? origin : allowedOrigin
}

export default {
  async fetch(request, env) {
    const origin = getOrigin(request)
    if (request.method === 'OPTIONS') return new Response(null, { headers: {
      'access-control-allow-origin': origin,
      'access-control-allow-headers': 'content-type, x-user-email, x-user-name',
      'access-control-allow-methods': 'GET, POST, OPTIONS',
    } })

    const url = new URL(request.url)
    if (url.pathname !== '/messages') return json({ error: 'Not found' }, 404, origin)

    const channel = url.searchParams.get('channel') || 'general'
    if (request.method === 'GET') {
      const result = await env.DB.prepare(
        'SELECT id, channel, user_email AS userEmail, user_name AS userName, text, created_at AS createdAt FROM messages WHERE channel = ? ORDER BY created_at ASC LIMIT 100',
      ).bind(channel).all()
      return json(result.results, 200, origin)
    }

    if (request.method === 'POST') {
      const email = request.headers.get('X-User-Email')?.trim().toLowerCase()
      const name = request.headers.get('X-User-Name')?.trim() || email?.split('@')[0]
      const body = await request.json()
      const text = typeof body.text === 'string' ? body.text.trim() : ''
      if (!email || !text || text.length > 2000) return json({ error: 'Email and a message under 2,000 characters are required.' }, 400, origin)
      const id = crypto.randomUUID()
      await env.DB.prepare('INSERT INTO messages (id, channel, user_email, user_name, text) VALUES (?, ?, ?, ?, ?)').bind(id, channel, email, name, text).run()
      return json({ id, channel, userEmail: email, userName: name, text, createdAt: new Date().toISOString() }, 201, origin)
    }

    return json({ error: 'Method not allowed' }, 405, origin)
  },
}
