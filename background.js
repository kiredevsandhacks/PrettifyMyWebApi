async function prettifyWebApi(request, sender, sendResponse) {
  chrome.tabs.query({ active: true, currentWindow: true }, async function (tabs) {
    if (sender.id === 'bhfdhnhbnamllpiaaapodfmoicgbbcmc' && request.action === 'prettifyWebApi' && tabs[0].id === sender.tab.id) {
      await chrome.scripting.executeScript({
        target: { tabId: sender.tab.id },
        files: ['prettifyWebApi.js']
      });
    }
  });
}

chrome.runtime.onInstalled.addListener(function (details) { 
  chrome.runtime.onMessage.addListener(prettifyWebApi);

  chrome.action.onClicked.addListener(async function (tab) {
    await chrome.scripting.executeScript({
      target: { tabId: tab.id },
      files: ['content.js']
    });
  });

});