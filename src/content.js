(function () {
    // This file is injected on every click on the extension button, but only needs to set up its listener once.
    if (window.initialized) {
        return;
    }

    window.initialized = true;

    const api = globalThis.browser ?? globalThis.chrome;

    function send(message) {
        Promise.resolve(api.runtime.sendMessage(message)).catch(err => console.error('PrettifyMyWebApi:', err));
    }

    // eventHandler.js is injected into the page by background.js (world: MAIN) and talks to us through window.postMessage
    function prettifyWebApiHandler(event) {
        if (event.origin !== window.origin || !event.data) {
            return;
        }

        if (event.data.action === 'prettifyWebApi') {
            send({ action: "prettifyWebApi" });
        } else if (event.data.action === 'openInWebApi') {
            if (event.data.url.startsWith('https://')) {
                send({ action: "openInWebApi", url: event.data.url });
            }
        } else if (event.data.action === 'openFlowInWebApi') {
            if (event.data.url.startsWith('https://')) {
                send({ action: "openFlowInWebApi", url: event.data.url });
            }
        }
    }

    window.addEventListener('message', prettifyWebApiHandler);
})();
