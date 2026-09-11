import { StrictMode, useEffect, useState } from 'react'
import { createRoot } from 'react-dom/client'
import { PublicClientApplication } from '@azure/msal-browser'
import './styles.css'

const microsoftClientId = import.meta.env.VITE_MICROSOFT_CLIENT_ID
const microsoftRedirectUri = `${window.location.origin}${import.meta.env.BASE_URL}`
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
const blockedWords = ['badword', 'idiot', 'stupid', "motherfucker", "bitch", "asshole", "dumbass", "dickhead", "faggot", "cunt", "slut", "whore", "nigger", "nigga", "retard", "retarded"]
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

function App() {
  const [signedIn, setSignedIn] = useState(false)
  const [channel, setChannel] = useState('general')
  const [messages, setMessages] = useState(initialMessages)
  const [draft, setDraft] = useState('')
  const [filtered, setFiltered] = useState(false)
  if (!signedIn) return <Login onLogin={() => setSignedIn(true)} />
  const send = (event) => { event.preventDefault(); if (!draft.trim()) return; const clean = filterText(draft.trim()); setMessages([...messages, { name: 'You', initials: 'YO', color: 'mint', time: 'Just now', text: clean }]); setFiltered(clean !== draft.trim()); setDraft('') }
  return <div className="app"><aside className="rail"><div className="server">ST</div><div className="rail-rule" /><button className="rail-selected">⌂</button><button>＋</button><div className="rail-grow" /><button>?</button></aside>
    <aside className="channels"><header><small>YOUR CAMPUS</small><strong>Lakeside Academy</strong><button>⌄</button></header><div className="channel-list"><label>CHANNELS <span>＋</span></label>{channels.map((item) => <button key={item} className={channel === item ? 'channel-active' : ''} onClick={() => setChannel(item)}><i>#</i>{item}</button>)}</div><div className="profile"><Avatar initials="YO" color="mint" online /><span>You<small>Online</small></span><button>⚙</button></div></aside>
    <main className="chat"><header className="chat-top"><div><span>#</span><strong>{channel}</strong><small>Keep it kind, keep it curious.</small></div><nav><button>⌕</button><button>♢</button><button>♧</button><i /><button>?</button></nav></header><section className="messages"><div className="welcome"><b>#</b><h2>Welcome to #{channel}</h2><p>This is the beginning of the #{channel} channel.</p></div>{messages.map((message, index) => <article key={`${message.time}-${index}`}><Avatar initials={message.initials} color={message.color} /><div><header><strong>{message.name}</strong>{message.name === 'Maya Chen' && <small>MOD</small>}<time>{message.time}</time></header><p>{message.text}</p></div></article>)}</section><div className="composer-wrap">{filtered && <div className="filter-message">✦ Your message was softened by the community filter.</div>}<form className="composer" onSubmit={send}><button type="button">＋</button><input value={draft} onChange={(event) => setDraft(event.target.value)} placeholder={`Message #${channel}`} /><button type="button">☺</button><button className="send" aria-label="Send">↑</button></form><small className="hint">Messages are automatically checked for respectful language.</small></div></main>
    <aside className="members"><header>MEMBERS — 1 <button>⌕</button></header><label>ONLINE — 1</label><div className="member"><Avatar initials="YO" color="mint" online /><span>You<small>Online</small></span></div></aside>
  </div>
}

createRoot(document.getElementById('root')).render(<StrictMode><App /></StrictMode>)
