// the Web Api of an environment is on the same origin as its forms and views
function isSameOrigin(url, origin) {
  try {
    return new URL(url).origin === origin;
  } catch {
    return false;
  }
}

// the Web Api of a Dataverse environment, e.g. https://contoso.crm4.dynamics.com/api/data/v9.2/workflows
function isDataverseOnlineUrl(url) {
  try {
    const parsedUrl = new URL(url);
    return parsedUrl.protocol === 'https:' && parsedUrl.hostname.endsWith('.dynamics.com') && parsedUrl.pathname.startsWith('/api/data/');
  } catch {
    return false;
  }
}

async function handler(request, sender, sendResponse) {
  chrome.tabs.query({ active: true, currentWindow: true }, async function (tabs) {
    if (request.action === 'prettifyWebApi' && tabs[0].id === sender.tab.id) {

      await chrome.scripting.executeScript({
        target: { tabId: sender.tab.id },
        files: ['prettifyWebApi.js']
      });
    } else if (request.action === 'openInWebApi' && tabs[0].id === sender.tab.id && isSameOrigin(request.url, sender.origin)) {
      chrome.tabs.create({ url: request.url, active: true });
    } else if (request.action === 'openFlowInWebApi' && tabs[0].id === sender.tab.id && isDataverseOnlineUrl(request.url)) {
      chrome.tabs.create({ url: request.url, active: true });
    }
  });
}

chrome.action.onClicked.addListener(async function (tab) {
  chrome.runtime.onMessage.removeListener(handler);
  chrome.runtime.onMessage.addListener(handler);

  await chrome.scripting.executeScript({
    target: { tabId: tab.id },
    files: ['content.js']
  });
});

chrome.commands.onCommand.addListener((shortcut) => {
  const isProdMode = 'update_url' in chrome.runtime.getManifest();
  if (!isProdMode) {
    if (shortcut.includes("+M")) {
      chrome.runtime.reload();
    }
  }
});