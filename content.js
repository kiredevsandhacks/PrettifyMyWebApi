window.scriptTag = document.createElement('script');
scriptTag.src = chrome.runtime.getURL('eventHandler.js');
scriptTag.type = "text/javascript";

document.head.appendChild(scriptTag);

function handler(event) {
    if (event.source === window && event.data.appId === 'bhfdhnhbnamllpiaaapodfmoicgbbcmc') {
        if (event.data.action === 'prettifyWebApi') {
            chrome.runtime.sendMessage("bhfdhnhbnamllpiaaapodfmoicgbbcmc", event.data);
        } else if (event.data.action === 'openInWebApi') {
            if (event.data.url.startsWith('https://')) {
                window.open(event.data.url);
            }
        }
    }
}

if (!window.initialized) {
    window.addEventListener('message', handler);
}

window.initialized = true;