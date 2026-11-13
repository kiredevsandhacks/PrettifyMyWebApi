// Works in Chrome (service worker) and Firefox (event page).
// Firefox exposes the promise based `browser` namespace, Chrome only has `chrome`.
const api = globalThis.browser ?? globalThis.chrome;

const WEB_API_PATH = /\/api\/data\/v[0-9][0-9]?(\.[0-9])?\//;

// Firefox only (needs the webRequest + webRequestBlocking permissions, see firefox/manifest.json).
// Firefox renders every top level 'application/json' response in its own JSON viewer and sets a strict
// Content-Security-Policy on that page ("default-src 'none'; script-src resource:"). That replaces the
// page we want to prettify and blocks the styles and scripts we add to it. Serving the Web Api
// response as plain text gives us the same page Chrome shows, a body with a single <pre>.
function serveWebApiAsPlainText(details) {
  let pathname;
  try {
    pathname = new URL(details.url).pathname;
  } catch {
    return {};
  }

  if (!WEB_API_PATH.test(pathname)) {
    return {};
  }

  const contentType = (details.responseHeaders || []).find(h => h.name.toLowerCase() === 'content-type');

  if (!contentType || !/json/i.test(contentType.value || '')) {
    return {};
  }

  contentType.value = 'text/plain; charset=utf-8';

  return { responseHeaders: details.responseHeaders };
}

if (api.webRequest?.onHeadersReceived) {
  api.webRequest.onHeadersReceived.addListener(
    serveWebApiAsPlainText,
    { urls: ['*://*.dynamics.com/*'], types: ['main_frame'] },
    ['blocking', 'responseHeaders']
  );
}

// Firefox containers. A tab that is created without a cookieStoreId opens in the default container,
// where the user is not logged in. Chrome has no cookieStoreId, so this returns nothing there.
// The private browsing store cannot be set by id, a tab created in the window of a private tab is private already.
function newTabOptions(sourceTab, url) {
  const options = { url: url, active: true, windowId: sourceTab.windowId };

  if (sourceTab.cookieStoreId && sourceTab.cookieStoreId.startsWith('firefox-container-')) {
    options.cookieStoreId = sourceTab.cookieStoreId;
  }

  return options;
}

async function handleMessage(request, sender) {
  if (!sender.tab) {
    return;
  }

  const [activeTab] = await api.tabs.query({ active: true, currentWindow: true });

  if (!activeTab || activeTab.id !== sender.tab.id) {
    return;
  }

  if (request.action === 'prettifyWebApi') {
    await api.scripting.executeScript({
      target: { tabId: sender.tab.id },
      files: ['prettifyWebApi.js']
    });
  } else if (request.action === 'openInWebApi' || request.action === 'openFlowInWebApi') {
    await api.tabs.create(newTabOptions(activeTab, request.url));
  }
}

// Registered at the top level, so the listener is there when the event page or service worker wakes up for a message.
api.runtime.onMessage.addListener((request, sender) => {
  handleMessage(request, sender).catch(err => console.error('PrettifyMyWebApi:', err));
});

api.action.onClicked.addListener(async function (tab) {
  try {
    // content.js listens for the messages that eventHandler.js posts
    await api.scripting.executeScript({
      target: { tabId: tab.id },
      files: ['content.js']
    });

    // eventHandler.js needs the page's own objects (Xrm), so it runs in the page itself.
    // Injecting it this way is not affected by the Content-Security-Policy of the page.
    await api.scripting.executeScript({
      target: { tabId: tab.id },
      world: 'MAIN',
      files: ['eventHandler.js']
    });
  } catch (err) {
    console.error('PrettifyMyWebApi:', err);
  }
});

// Development shortcut (Ctrl+M reloads the extension). The Firefox manifest has no commands.
api.commands?.onCommand.addListener((shortcut) => {
  const isProdMode = 'update_url' in api.runtime.getManifest();
  if (!isProdMode) {
    if (shortcut.includes("+M")) {
      api.runtime.reload();
    }
  }
});
