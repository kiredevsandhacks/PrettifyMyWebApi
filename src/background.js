async function handler(request, sender, sendResponse) {
  chrome.tabs.query({ active: true, currentWindow: true }, async function (tabs) {
    if (request.action === 'prettifyWebApi' && tabs[0].id === sender.tab.id) {
      await chrome.scripting.executeScript({
        target: { tabId: sender.tab.id },
        files: ['prettifyWebApi.js']
      });
    } else if (request.action === 'openInWebApi' && tabs[0].id === sender.tab.id) {
      chrome.tabs.create({ url: request.url, active: true });
    } else if (request.action === 'openFlowInWebApi' && tabs[0].id === sender.tab.id) {
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
