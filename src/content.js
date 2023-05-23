window.scriptTag = document.createElement('script');
scriptTag.src = chrome.runtime.getURL('eventHandler.js');
scriptTag.type = "text/javascript";

document.head.appendChild(scriptTag);

function handler(event) {
    if (event.origin === window.origin) {
        if (event.data.action === 'prettifyWebApi') {
            chrome.runtime.sendMessage({ action: "prettifyWebApi" });
        } else if (event.data.action === 'openInWebApi') {
            if (event.data.url.startsWith('https://')) {
                chrome.runtime.sendMessage({ action: "openInWebApi", url: event.data.url });
            }
        }
    }
}

if (!window.initialized) {
    window.addEventListener('message', handler);
}

window.initialized = true;