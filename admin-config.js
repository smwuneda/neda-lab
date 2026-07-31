// NEDA Lab administrator configuration
// 1) Paste the Firebase web app configuration from Firebase Console.
// 2) Paste the same Apps Script URL used in neda-config.js.
window.NEDA_ADMIN_CONFIG = {
  firebase: {
    apiKey: '',
    authDomain: '',
    projectId: '',
    storageBucket: '',
    messagingSenderId: '',
    appId: ''
  },
  appsScriptUrl: '',
  // This is a UI check only. The Apps Script backend also checks its own allowlist.
  allowedEmails: [
    'YOUR_ADMIN_EMAIL@gmail.com'
  ]
};
