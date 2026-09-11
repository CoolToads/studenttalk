const allowedOrigin = 'https://cooltoads.github.io'

function getOrigin(request) {
  const origin = request.headers.get('Origin')

  if (origin === allowedOrigin) return origin
  if (origin && /^http:\/\/localhost:\d+$/.test(origin)) return origin

  return allowedOrigin
}

function corsHeaders(origin) {
  return {
    'Access-Control-Allow-Origin': origin,
    'Access-Control-Allow-Headers': 'Content-Type, X-User-Email, X-User-Name',
    'Access-Control-Allow-Methods': 'GET, POST, PUT, DELETE, OPTIONS',
    'Vary': 'Origin',
  }
}

function json(data, status = 200, origin = allowedOrigin) {
  return new Response(JSON.stringify(data), {
    status,
    headers: {
      'Content-Type': 'application/json; charset=utf-8',
      ...corsHeaders(origin),
    },
  })
}

export default {
  async fetch(request, env) {
    const origin = getOrigin(request)

    // CORS preflight
    if (request.method === 'OPTIONS') {
      return new Response(null, {
        status: 204,
        headers: corsHeaders(origin),
      })
    }

    const url = new URL(request.url)

    // -------------------------
    // PROFILE
    // -------------------------

    if (url.pathname === '/profile') {
      const email = request.headers
        .get('X-User-Email')
        ?.trim()
        .toLowerCase()

      if (!email) {
        return json(
          { error: 'A signed-in email is required.' },
          401,
          origin
        )
      }

      if (request.method === 'GET') {
        const profile = await env.DB
          .prepare(
            'SELECT email, username FROM profiles WHERE email = ?'
          )
          .bind(email)
          .first()

        if (!profile) {
          return json(
            { error: 'Profile not found.' },
            404,
            origin
          )
        }

        return json(profile, 200, origin)
      }

      if (request.method === 'PUT') {
        let body

        try {
          body = await request.json()
        } catch {
          return json(
            { error: 'Invalid JSON body.' },
            400,
            origin
          )
        }

        const username =
          typeof body.username === 'string'
            ? body.username.trim()
            : ''

        if (!/^[A-Za-z0-9_ -]{3,24}$/.test(username)) {
          return json(
            {
              error:
                'Username must be 3-24 characters using letters, numbers, spaces, hyphens, or underscores.',
            },
            400,
            origin
          )
        }

        try {
          await env.DB
            .prepare(`
              INSERT INTO profiles (email, username)
              VALUES (?, ?)
              ON CONFLICT(email)
              DO UPDATE SET
                username = excluded.username,
                updated_at = CURRENT_TIMESTAMP
            `)
            .bind(email, username)
            .run()

          return json(
            {
              email,
              username,
            },
            200,
            origin
          )
        } catch (error) {
          if (String(error).toLowerCase().includes('unique')) {
            return json(
              { error: 'That username is already taken.' },
              409,
              origin
            )
          }

          console.error(error)

          return json(
            { error: 'Failed to save profile.' },
            500,
            origin
          )
        }
      }

      return json(
        { error: 'Method not allowed' },
        405,
        origin
      )
    }

    // -------------------------
    // MESSAGES
    // -------------------------

    if (url.pathname !== '/messages') {
      return json(
        { error: 'Not found' },
        404,
        origin
      )
    }

    const channel =
      url.searchParams.get('channel')?.trim() || 'general'

    // GET messages
    if (request.method === 'GET') {
      const result = await env.DB
        .prepare(`
          SELECT
            id,
            channel,
            user_email AS userEmail,
            user_name AS userName,
            text,
            created_at AS createdAt
          FROM messages
          WHERE channel = ?
          ORDER BY created_at ASC
          LIMIT 100
        `)
        .bind(channel)
        .all()

      return json(
        result.results,
        200,
        origin
      )
    }

    // POST message
    if (request.method === 'POST') {
      const email = request.headers
        .get('X-User-Email')
        ?.trim()
        .toLowerCase()

      const name =
        request.headers.get('X-User-Name')?.trim() ||
        email?.split('@')[0] ||
        ''

      if (!email) {
        return json(
          { error: 'A signed-in email is required.' },
          401,
          origin
        )
      }

      let body

      try {
        body = await request.json()
      } catch {
        return json(
          { error: 'Invalid JSON body.' },
          400,
          origin
        )
      }

      const text =
        typeof body.text === 'string'
          ? body.text.trim()
          : ''

      if (!text) {
        return json(
          { error: 'A message is required.' },
          400,
          origin
        )
      }

      if (text.length > 2000) {
        return json(
          {
            error:
              'Messages cannot be longer than 2,000 characters.',
          },
          400,
          origin
        )
      }

      const id = crypto.randomUUID()
      const createdAt = new Date().toISOString()

      try {
        await env.DB
          .prepare(`
            INSERT INTO messages
              (id, channel, user_email, user_name, text)
            VALUES (?, ?, ?, ?, ?)
          `)
          .bind(
            id,
            channel,
            email,
            name,
            text
          )
          .run()

        return json(
          {
            id,
            channel,
            userEmail: email,
            userName: name,
            text,
            createdAt,
          },
          201,
          origin
        )
      } catch (error) {
        console.error(error)

        return json(
          { error: 'Failed to send message.' },
          500,
          origin
        )
      }
    }

    // DELETE message
    if (request.method === 'DELETE') {
      const email = request.headers
        .get('X-User-Email')
        ?.trim()
        .toLowerCase()

      const id =
        url.searchParams.get('id')?.trim()

      if (!email || !id) {
        return json(
          {
            error:
              'A signed-in email and message ID are required.',
          },
          400,
          origin
        )
      }

      const result = await env.DB
        .prepare(
          'DELETE FROM messages WHERE id = ? AND user_email = ?'
        )
        .bind(id, email)
        .run()

      if (!result.meta.changes) {
        return json(
          {
            error:
              'Message not found or you are not allowed to delete it.',
          },
          404,
          origin
        )
      }

      return json(
        {
          deleted: true,
          id,
        },
        200,
        origin
      )
    }

    return json(
      { error: 'Method not allowed' },
      405,
      origin
    )
  },
}