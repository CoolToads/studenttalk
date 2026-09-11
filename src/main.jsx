
/* Why are you in the source code? */

import { StrictMode, useEffect, useState } from 'react'
import { createRoot } from 'react-dom/client'
import { PublicClientApplication } from '@azure/msal-browser'
import './styles.css'

const microsoftClientId = import.meta.env.VITE_MICROSOFT_CLIENT_ID
const microsoftRedirectUri = `${window.location.origin}${import.meta.env.BASE_URL}`
const apiUrl = import.meta.env.VITE_API_URL?.replace(/\/$/, '')
const msalInstance = microsoftClientId ? new PublicClientApplication({
  auth: {
    clientId: microsoftClientId,
    authority: 'https://login.microsoftonline.com/common',
    redirectUri: microsoftRedirectUri,
  },
  cache: { cacheLocation: 'sessionStorage' },
}) : null

const channels = ['general', 'homework-help', 'study-groups', 'off-topic']
const initialMessages = []
const blockedWords = ['badword', 'idiot', 'stupid', "motherfucker", "bitch", "asshole", "dumbass", "dickhead", "faggot", "cunt", "slut", "whore", "nigger", "nigga", "retard", "retarded", "fuck", "shit", "bastard", "ass", "cock", "pussy", "dick", "fag", "twat", "wanker", "bollocks", "arsehole", "bugger", "bloody", "bollocking", "bollock", "arse", "prick", "tosser", "knobhead"]
const filterText = (text) => blockedWords.reduce((result, word) => result.replace(new RegExp(word, 'gi'), '•••'), text)

function MicrosoftLogo() { return <span className="ms-logo"><i /><i /><i /><i /></span> }
function Avatar({ initials, color, online = false }) { return <div className={`avatar ${color}`}>{initials}{online && <span className="online" />}</div> }

function Login({ onLogin }) {
  const [loginError, setLoginError] = useState('')
  const [isLoading, setIsLoading] = useState(Boolean(msalInstance))
  const [isSigningIn, setIsSigningIn] = useState(false)

  useEffect(() => {
    if (!msalInstance) {
      setIsLoading(false)
      return
    }
    let mounted = true
    const finishRedirectLogin = async () => {
      try {
        await msalInstance.initialize()
        const response = await msalInstance.handleRedirectPromise()
        const account = response?.account ?? msalInstance.getAllAccounts()[0]
        if (!account) return
        msalInstance.setActiveAccount(account)
        await msalInstance.acquireTokenSilent({ scopes: ['User.Read'], account })
        if (mounted) onLogin(account)
      } catch (error) {
        if (mounted) setLoginError(error instanceof Error ? error.message : 'Microsoft sign-in could not be completed.')
      } finally {
        if (mounted) setIsLoading(false)
      }
    }
    finishRedirectLogin()
    return () => { mounted = false }
  }, [onLogin])

  const startMicrosoftLogin = async () => {
    if (!msalInstance) {
      setLoginError('Microsoft OAuth is not configured yet. Add VITE_MICROSOFT_CLIENT_ID to enable school sign-in.')
      return
    }
    setIsSigningIn(true)
    setLoginError('')
    try {
      await msalInstance.initialize()
      await msalInstance.loginRedirect({ scopes: ['openid', 'profile', 'email', 'User.Read'] })
    } catch (error) {
      setIsSigningIn(false)
      setLoginError(error instanceof Error ? error.message : 'Microsoft sign-in could not be started.')
    }
  }
  return <main className="login-page"><div className="glow glow-one" /><div className="glow glow-two" /><section className="login-card">
    <div className="brand"><strong>ST</strong><span>Student Talk</span></div>
    <div className="login-copy"><small>A CALMER PLACE TO CONNECT</small><h1>Find your people.<br /><em>Build your space.</em></h1><p>A focused community for students to collaborate, share ideas, and make school feel a little more connected. Brought to you by JPRG Hub</p></div>
    <button className="microsoft-login" onClick={msalInstance ? startMicrosoftLogin : onLogin} disabled={isLoading || isSigningIn}><MicrosoftLogo /> {isLoading ? 'Checking your session...' : isSigningIn ? 'Redirecting to Microsoft...' : 'Continue with Microsoft'} <b>→</b></button>
    {loginError && <p className="login-error">{loginError}</p>}
    <p className="login-note">Sign in with your school Microsoft account</p>
    <div className="trust">✓ Protected by your school account <span>·</span> Kind by design</div>
  </section><footer>Student Talk <span>•</span> Built for better conversations</footer></main>
}

function ProfileSetup({ email, onComplete }) {
  const [username, setUsername] = useState('')
  const [error, setError] = useState('')
  const [saving, setSaving] = useState(false)

  const saveProfile = async (event) => {
    event.preventDefault()
    const cleanUsername = username.trim()
    if (!/^[A-Za-z0-9_ -]{3,24}$/.test(cleanUsername)) {
      setError('Use 3-24 letters, numbers, spaces, hyphens, or underscores.')
      return
    }
    setSaving(true)
    setError('')
    if (!apiUrl) {
      localStorage.setItem(`student-talk-username:${email}`, cleanUsername)
      onComplete({ email, username: cleanUsername })
      return
    }
    try {
      const response = await fetch(`${apiUrl}/profile`, { method: 'PUT', headers: { 'content-type': 'application/json', 'x-user-email': email }, body: JSON.stringify({ username: cleanUsername }) })
      const result = await response.json()
      if (!response.ok) throw new Error(result.error || 'Could not save your username.')
      onComplete(result)
    } catch (requestError) {
      setError(requestError instanceof Error ? requestError.message : 'Could not save your username.')
      setSaving(false)
    }
  }

  return <main className="login-page"><div className="glow glow-one" /><div className="glow glow-two" /><section className="login-card profile-setup"><div className="brand"><strong>ST</strong><span>Student Talk</span></div><div className="login-copy"><small>ONE LAST THING</small><h1>Choose your<br /><em>display name.</em></h1><p>This is how other students will see you in conversations. You can change it later in Settings.</p></div><form onSubmit={saveProfile}><label className="profile-label" htmlFor="username">Username</label><input id="username" className="profile-input" value={username} onChange={(event) => setUsername(event.target.value)} placeholder="e.g. jordan_w" autoComplete="nickname" autoFocus /><small className="profile-email">{email}</small><button className="microsoft-login" disabled={saving}>{saving ? 'Saving profile...' : 'Continue'} <b>→</b></button>{error && <p className="login-error">{error}</p>}</form></section><footer>Student Talk <span>•</span> Your space, your name</footer></main>
}

function App() {
  const [signedIn, setSignedIn] = useState(false)
  const [account, setAccount] = useState(null)
  const [profile, setProfile] = useState(null)
  const [profileLoading, setProfileLoading] = useState(false)
  const [profileError, setProfileError] = useState('')
  const [channel, setChannel] = useState('general')
  const [messages, setMessages] = useState(initialMessages)
  const [draft, setDraft] = useState('')
  const [filtered, setFiltered] = useState(false)
  const handleLogin = (nextAccount) => { setAccount(nextAccount); setSignedIn(true) }
  const email = account?.username || 'local-user'

  useEffect(() => {
    if (!signedIn || profile) return
    const localUsername = localStorage.getItem(`student-talk-username:${email}`)
    if (localUsername) { setProfile({ email, username: localUsername }); return }
    if (!apiUrl) return
    let active = true
    setProfileLoading(true)
    fetch(`${apiUrl}/profile`, { headers: { 'x-user-email': email } }).then(async (response) => {
      const result = await response.json()
      if (response.status === 404) return null
      if (!response.ok) throw new Error(result.error || 'Could not load your profile.')
      return result
    }).then((result) => { if (active && result) setProfile(result) }).catch((error) => { if (active) setProfileError(error instanceof Error ? error.message : 'Could not load your profile.') }).finally(() => { if (active) setProfileLoading(false) })
    return () => { active = false }
  }, [signedIn, profile, email])

  if (!signedIn) return <Login onLogin={handleLogin} />
  if (profileLoading) return <main className="login-page"><section className="login-card loading-card"><div className="brand"><strong>ST</strong><span>Student Talk</span></div><p>Loading your profile...</p></section></main>
  if (!profile) return <ProfileSetup email={email} onComplete={setProfile} />
  const send = (event) => { event.preventDefault(); if (!draft.trim()) return; const clean = filterText(draft.trim()); setMessages([...messages, { id: crypto.randomUUID(), name: profile.username, userEmail: email, initials: profile.username.slice(0, 2).toUpperCase(), color: 'mint', time: 'Just now', text: clean }]); setFiltered(clean !== draft.trim()); setDraft('') }
  const deleteMessage = async (message) => {
    if (!window.confirm('Delete this message?')) return
    if (apiUrl && message.id) {
      const response = await fetch(`${apiUrl}/messages?id=${encodeURIComponent(message.id)}`, { method: 'DELETE', headers: { 'x-user-email': email } })
      if (!response.ok) return
    }
    setMessages((current) => current.filter((item) => item.id !== message.id))
  }
  return <div className="app"><aside className="rail"><div className="server">ST</div><div className="rail-rule" /><button className="rail-selected">⌂</button><button>＋</button><div className="rail-grow" /><button>?</button></aside>
    <aside className="channels"><header><small>YOUR CAMPUS</small><strong>Lakeside Academy</strong><button>⌄</button></header><div className="channel-list"><label>CHANNELS <span>＋</span></label>{channels.map((item) => <button key={item} className={channel === item ? 'channel-active' : ''} onClick={() => setChannel(item)}><i>#</i>{item}</button>)}</div><div className="profile"><Avatar initials={profile.username.slice(0, 2).toUpperCase()} color="mint" online /><span>{profile.username}<small>Online</small></span><button>⚙</button></div></aside>
    <main className="chat"><header className="chat-top"><div><span>#</span><strong>{channel}</strong><small>Keep it kind, keep it curious.</small></div><nav><button>⌕</button><button>♢</button><button>♧</button><i /><button>?</button></nav></header><section className="messages"><div className="welcome"><b>#</b><h2>Welcome to #{channel}</h2><p>This is the beginning of the #{channel} channel.</p></div>{messages.map((message, index) => <article key={message.id || `${message.time}-${index}`}><Avatar initials={message.initials} color={message.color} /><div><header><strong>{message.name}</strong>{message.name === 'Maya Chen' && <small>MOD</small>}<time>{message.time}</time>{message.userEmail === email && <button className="delete-message" onClick={() => deleteMessage(message)} title="Delete message" aria-label="Delete message">⌫</button>}</header><p>{message.text}</p></div></article>)}</section><div className="composer-wrap">{filtered && <div className="filter-message">✦ Your message was softened by the community filter.</div>}<form className="composer" onSubmit={send}><button type="button">＋</button><input value={draft} onChange={(event) => setDraft(event.target.value)} placeholder={`Message #${channel}`} /><button type="button">☺</button><button className="send" aria-label="Send">↑</button></form><small className="hint">Messages are automatically checked for respectful language.</small></div></main>
    <aside className="members"><header>MEMBERS — 1 <button>⌕</button></header><label>ONLINE — 1</label><div className="member"><Avatar initials={profile.username.slice(0, 2).toUpperCase()} color="mint" online /><span>{profile.username}<small>{profileError || 'Online'}</small></span></div></aside>
  </div>
}

createRoot(document.getElementById('root')).render(<StrictMode><App /></StrictMode>)
