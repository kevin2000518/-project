(() => {
  const installButton = document.querySelector('#installApp');
  const help = document.querySelector('#installHelp');
  let installPrompt;
  const installed = () => window.matchMedia('(display-mode: standalone)').matches || navigator.standalone === true;
  if (installed()) {
    installButton.hidden = true;
    help.textContent = '已以 App 模式開啟金煎。';
  }
  window.addEventListener('beforeinstallprompt', event => {
    event.preventDefault();
    installPrompt = event;
    if (!installed()) installButton.hidden = false;
  });
  installButton.addEventListener('click', async () => {
    if (!installPrompt) return;
    const prompt = installPrompt;
    installPrompt = null;
    installButton.hidden = true;
    try {
      await prompt.prompt();
      await prompt.userChoice;
    } catch (error) {
      console.warn('Unable to show install prompt', error);
    }
  });
  window.addEventListener('appinstalled', () => {
    installPrompt = null;
    installButton.hidden = true;
    help.textContent = '金煎已安裝至您的裝置。';
  });
  if ('serviceWorker' in navigator) {
    window.addEventListener('load', () => {
      navigator.serviceWorker.register('./sw.js', {scope: './'})
        .catch(error => console.warn('PWA offline setup failed', error));
    });
  }
})();
