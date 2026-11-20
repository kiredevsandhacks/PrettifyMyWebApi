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

// the Web Api of a Dataverse environment, e.g. https://contoso.crm4.dynamics.com/api/data/v9.2/workflows
function isDataverseOnlineUrl(url) {
    try {
        const parsedUrl = new URL(url);
        return parsedUrl.protocol === 'https:' && parsedUrl.hostname.endsWith('.dynamics.com') && parsedUrl.pathname.startsWith('/api/data/');
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
            // a flow is opened from Power Automate in the Web Api of its Dataverse environment, which is on another origin
            if (isDataverseOnlineUrl(event.data.url)) {
                chrome.runtime.sendMessage({ action: "openFlowInWebApi", url: event.data.url });
            }
        }
    }
}

if (!window.initialized) {
    window.addEventListener('message', prettifyWebApiHandler);
}

window.initialized = true;