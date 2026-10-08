import { isLaunchedByHaven, renderHavenAppLandingPage } from 'mindoodb-app-sdk'

// Opened directly (a shared link, a bookmark) there is no Haven to talk to: show what
// the app is and a button that installs it, instead of booting into a connection error.
// `?standalone` opens the editor without Haven (tests, demos).
if (!isLaunchedByHaven() && !new URLSearchParams(window.location.search).has('standalone')) {
  void renderHavenAppLandingPage()
} else {
  void import('./main')
}
