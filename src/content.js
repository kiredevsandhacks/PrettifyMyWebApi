window.scriptTag = document.createElement('script');
scriptTag.src = chrome.runtime.getURL('eventHandler.js');
scriptTag.type = "text/javascript";

document.head.appendChild(scriptTag);

function isOnThisOrigin(url) {
    try {
        return typeof url === 'string' && new URL(url).origin === window.location.origin;
    } catch {
        return false;
    }
}

function prettifyWebApiHandler(event) {
    if (event.origin === window.origin) {
        if (event.data.action === 'prettifyWebApi') {
            chrome.runtime.sendMessage({ action: "prettifyWebApi" });
        } else if (event.data.action === 'openInWebApi') {
            // the Web Api of this environment is on the same origin as the form or view
            if (isOnThisOrigin(event.data.url)) {
                chrome.runtime.sendMessage({ action: "openInWebApi", url: event.data.url });
            }
        } else if (event.data.action === 'openFlowInWebApi') {
            if (event.data.url.startsWith('https://')) {
                chrome.runtime.sendMessage({ action: "openFlowInWebApi", url: event.data.url });
            }
        }
    }
}

if (!window.initialized) {
    window.addEventListener('message', prettifyWebApiHandler);
}

window.initialized = true;