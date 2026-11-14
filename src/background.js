// Works in Chrome (service worker) and Firefox (event page).
// Firefox exposes the promise based `browser` namespace, Chrome only has `chrome`.
const api = globalThis.browser ?? globalThis.chrome;

const WEB_API_PATH = /\/api\/data\/v[0-9][0-9]?(\.[0-9])?\//;

// The extension itself puts one of these in the url of the pages it prettifies
const PRETTIFY_HASHES = ['#p', '#pf', '#pr'];

// How long a tab that was reloaded by the extension is treated as a prettify page, and the minimum time between two such reloads
const FORCED_RELOAD_MS = 10000;

// tabId -> time of the last reload as plain text, see reloadAsPlainText
const forcedReloads = new Map();

// Firefox only, needs the webRequest + webRequestBlocking permissions (see firefox/manifest.json)
const canServeAsPlainText = !!api.webRequest?.onHeadersReceived;

function isWebApiUrl(urlString) {
  try {
    const url = new URL(urlString);
    return url.hostname.endsWith('.dynamics.com') && WEB_API_PATH.test(url.pathname);
  } catch {
    return false;
  }
}

// Firefox renders every top level 'application/json' response in its own JSON viewer and sets a strict
// Content-Security-Policy on that page ("default-src 'none'; script-src resource:"). That replaces the
// page we want to prettify and blocks the styles and scripts we add to it. For the pages that we prettify
// (a Web Api url that ends in #p, #pf or #pr) the response is served as plain text instead, which gives us
// the same page Chrome shows, a body with a single <pre>.
// All other Web Api pages, and all fetch and XHR calls, keep their original content type.
function serveWebApiAsPlainText(details) {
  if (!isWebApiUrl(details.url)) {
    return {};
  }

  const forcedAt = forcedReloads.get(details.tabId);
  const forced = forcedAt !== undefined && Date.now() - forcedAt < FORCED_RELOAD_MS;

  if (!PRETTIFY_HASHES.includes(new URL(details.url).hash) && !forced) {
    return {};
  }

  const contentType = (details.responseHeaders || []).find(h => h.name.toLowerCase() === 'content-type');

  if (!contentType || !/json/i.test(contentType.value || '')) {
    return {};
  }

  contentType.value = 'text/plain; charset=utf-8';

  return { responseHeaders: details.responseHeaders };
}

// Reloads a Web Api tab with a #p url so the response is served as plain text, for pages that are not prettified yet.
// Returns false when it did not reload, because the tab is not a Web Api page or was reloaded a moment ago (no reload loops).
async function reloadAsPlainText(tab) {
  if (!tab || !isWebApiUrl(tab.url)) {
    return false;
  }

  const last = forcedReloads.get(tab.id);

  if (last !== undefined && Date.now() - last < FORCED_RELOAD_MS) {
    return false;
  }

  forcedReloads.set(tab.id, Date.now());

  const url = new URL(tab.url);

  if (!PRETTIFY_HASHES.includes(url.hash)) {
    url.hash = '#p';
  }

  if (url.href !== tab.url) {
    await api.tabs.update(tab.id, { url: url.href });
  }

  await api.tabs.reload(tab.id);

  return true;
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
  if (request.action === 'reloadAsPlainText') {
    // sent by prettifyWebApi.js when it finds the Firefox JSON viewer, the answer tells it whether a reload is on its way
    return reloadAsPlainText(sender.tab).catch(err => { console.error('PrettifyMyWebApi:', err); return false; });
  }

  handleMessage(request, sender).catch(err => console.error('PrettifyMyWebApi:', err));
});

api.action.onClicked.addListener(async function (tab) {
  try {
    // A Web Api page that is open now can be the Firefox JSON viewer, which the scripts cannot work on.
    if (canServeAsPlainText && isWebApiUrl(tab.url)) {
      await reloadAsPlainText(tab);
      return;
    }

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
