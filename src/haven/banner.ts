// A slim notice over the editor's status bar: "you are looking at an older
// revision / at Haven's time travel — read-only", optionally with a button.
const STYLE = `
.tg-banner{--tg-banner-action:#1f3a8a;--tg-banner-action-hover:#243c8f;--tg-banner-on-action:#ffffff;--tg-banner-focus:rgba(212,160,23,.38);position:fixed;left:50%;bottom:32px;transform:translateX(-50%);z-index:8500;display:flex;align-items:center;gap:12px;max-width:calc(100vw - 32px);padding:8px 14px;border-radius:4px;background:var(--surface);color:var(--text);border:1px solid var(--border-strong);box-shadow:0 6px 24px color-mix(in srgb,var(--text) 20%,transparent);font:13px/1.4 system-ui,sans-serif}
.tg-banner svg{flex:none;width:16px;height:16px}
.tg-banner button{all:unset;cursor:pointer;padding:4px 10px;border-radius:2px;background:var(--tg-banner-action);color:var(--tg-banner-on-action);font-weight:600;white-space:nowrap}
.tg-banner button:hover{background:var(--tg-banner-action-hover)}
.tg-banner button:focus-visible{outline:2px solid var(--tg-banner-focus);outline-offset:2px}
`
const CLOCK =
  '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.6" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true"><circle cx="12" cy="12" r="8.5"/><path d="M12 7.5V12l3 2"/></svg>'

let banner: HTMLElement | null = null

export function showBanner(text: string, action?: { label: string; run: () => void }): void {
  if (!document.getElementById('tg-banner-style')) {
    const style = document.createElement('style')
    style.id = 'tg-banner-style'
    style.textContent = STYLE
    document.head.append(style)
  }
  hideBanner()
  banner = document.createElement('div')
  banner.className = 'tg-banner'
  banner.setAttribute('role', 'status')
  banner.innerHTML = CLOCK
  const label = document.createElement('span')
  label.textContent = text
  banner.append(label)
  if (action) {
    const button = document.createElement('button')
    button.type = 'button'
    button.textContent = action.label
    button.addEventListener('click', action.run)
    banner.append(button)
  }
  document.body.append(banner)
}

export function hideBanner(): void {
  banner?.remove()
  banner = null
}
