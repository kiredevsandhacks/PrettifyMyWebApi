async function prettifyWebApi(request, sender, sendResponse) {
  chrome.tabs.query({ active: true, currentWindow: true }, async function (tabs) {
    if (request.action === 'prettifyWebApi' && tabs[0].id === sender.tab.id) {
      await chrome.scripting.executeScript({
        target: { tabId: sender.tab.id },
        files: ['prettifyWebApi.js']
      });
    }
  });
}

chrome.action.onClicked.addListener(async function (tab) {
  chrome.runtime.onMessage.removeListener(prettifyWebApi);
  chrome.runtime.onMessage.addListener(prettifyWebApi);

  await chrome.scripting.executeScript({
    target: { tabId: tab.id },
    files: ['content.js']
  });
});

