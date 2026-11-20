(async function () {
    const guidPattern = /^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$/i;
    const logicalNamePattern = /^[a-z0-9_]+$/i;

    async function odataFetch(url) {
        const response = await fetch(url, { headers: { 'Prefer': 'odata.include-annotations="*"', 'Cache-Control': 'no-cache' } });

        return await response.json();
    }

    // ids can be wrapped in braces, e.g. {00000000-0000-0000-0000-000000000000}
    function stripBraces(id) {
        return id.replace(/^\{(.*)\}$/, '$1');
    }

    // e.g. (00000000-0000-0000-0000-000000000000)
    function getRowKey(recordId) {
        const id = stripBraces(recordId);

        // e.g. a new row which is not saved yet
        if (!guidPattern.test(id)) {
            throw 'This row does not have an id yet, save it first';
        }

        return `(${id})`;
    }

    // e.g. ?savedQuery=00000000-0000-0000-0000-000000000000
    function getViewQuery(viewId, viewType) {
        const id = stripBraces(viewId);

        if (!guidPattern.test(id)) {
            throw 'unknown view id: ' + viewId;
        }

        if (viewType == 4230) {
            return `?userQuery=${id}`;
        } else if (viewType == 1039) {
            return `?savedQuery=${id}`;
        } else {
            throw 'unknown view type: ' + viewType;
        }
    }

    // The Web Api of the environment, from its base url, e.g. https://contoso.crm4.dynamics.com or https://my-env:1234/contoso,
    // and its full version, e.g. 9.2.24081.00191
    function buildApiUrl(clientUrl, fullVersion) {
        const versionArray = fullVersion.split('.');
        const version = versionArray[0] + '.' + versionArray[1];

        let apiUrl;
        try {
            // absolute, e.g. https://contoso.crm4.dynamics.com/api/data/v9.2/ or https://my-env:1234/contoso/api/data/v9.1/
            apiUrl = new URL(clientUrl.replace(/\/+$/, '') + `/api/data/v${version}/`, window.location.href).href;
        } catch {
            throw 'Could not determine the Web Api url of this environment from ' + clientUrl;
        }

        // the Web Api is on the origin of this page (protocol, host and port, for on-prem)
        if (new URL(apiUrl).origin !== window.location.origin) {
            throw 'The Web Api url of this environment (' + apiUrl + ') is not on the origin of this page (' + window.location.origin + ')';
        }

        return apiUrl;
    }

    // e.g. https://contoso.crm4.dynamics.com/api/data/v9.2/accounts
    async function getEntitySetUrl(apiUrl, entityLogicalName) {
        if (!logicalNamePattern.test(entityLogicalName)) {
            throw 'unknown table: ' + entityLogicalName;
        }

        const requestUrl = apiUrl + 'EntityDefinitions?$select=EntitySetName&$filter=(LogicalName eq %27' + entityLogicalName + '%27)';

        const result = await odataFetch(requestUrl)
        const pluralName = result.value[0].EntitySetName;

        return apiUrl + pluralName;
    }

    // Modern mode: the url of main.aspx says what is open, without needing Xrm, e.g.
    // a view: main.aspx?etn=account&viewid=...&viewType=1039
    // a row:  main.aspx?etn=account&id=...&pagetype=entityrecord
    // Returns null when the url does not say, then it is up to legacy mode.
    function getModernPage(urlObj) {
        const entityLogicalName = urlObj.searchParams.get('etn');
        const viewId = urlObj.searchParams.get('viewid');
        const viewType = urlObj.searchParams.get('viewType');
        const recordId = urlObj.searchParams.get('id');

        if (entityLogicalName && viewId && viewType) {
            return { entityLogicalName, viewId, viewType };
        }

        if (entityLogicalName && recordId) {
            return { entityLogicalName, recordId };
        }

        return null;
    }

    // The Web Api of the environment, from the url of main.aspx and the version the environment reports
    async function getModernApiUrl() {
        // Online and on-prem with IFD, main.aspx is at the root (the organization is in the host),
        // on-prem without IFD it comes after the organization, e.g. /contoso/main.aspx
        const pathMatch = /^(\/[^\/]+)?\/main\.aspx$/i.exec(window.location.pathname);

        if (pathMatch == null) {
            throw 'Could not determine the Web Api url of this environment from ' + window.location.pathname;
        }

        const clientUrl = window.location.origin + (pathMatch[1] ?? '');

        // v9.0 is available on every 9.x environment, online and on-prem
        let fullVersion;
        try {
            fullVersion = (await odataFetch(clientUrl + '/api/data/v9.0/RetrieveVersion()')).Version;
        } catch {
            // e.g. not a Dataverse environment
        }

        if (!/^\d+\.\d+\./.test(fullVersion ?? '')) {
            throw 'Could not determine the Web Api version of this environment from ' + clientUrl;
        }

        return buildApiUrl(clientUrl, fullVersion);
    }

    async function getModernWebApiUrl({ entityLogicalName, viewId, viewType, recordId }) {
        const query = viewId ? getViewQuery(viewId, viewType) : getRowKey(recordId);
        const apiUrl = await getModernApiUrl();
        const entitySetUrl = await getEntitySetUrl(apiUrl, entityLogicalName);

        return entitySetUrl + query;
    }

    // Legacy mode: the url does not say what is open, so the row comes from the form through Xrm.Page.
    // Views are not supported in legacy mode.
    async function getLegacyWebApiUrl() {
        const entity = window.Xrm.Page.data?.entity;

        if (!entity) {
            throw 'Please open a form or view to use PrettifyMyWebApi';
        }

        const rowKey = getRowKey(entity.getId());

        // documented to work online and on-prem, with the organization in the path or, with IFD, in the host
        const globalContext = window.Xrm.Utility.getGlobalContext();
        const apiUrl = buildApiUrl(globalContext.getClientUrl(), globalContext.getVersion());
        const entitySetUrl = await getEntitySetUrl(apiUrl, entity.getEntityName());

        return entitySetUrl + rowKey;
    }

    async function openInWebApi(getWebApiUrl) {
        let urlToOpen;
        try {
            urlToOpen = await getWebApiUrl();
        } catch (err) {
            alert(err);
            return;
        }

        urlToOpen += '#p'; // add the secret sauce
        window.postMessage({ action: 'openInWebApi', url: urlToOpen });
    }

    function getDataverseUrl() {
        const currentEnvironmentId = location.href.split('/environments/').pop().split('?')[0].split('/')[0];

        for (let i = 0; i < localStorage.length; i++) {
            const value = localStorage.getItem(localStorage.key(i));

            try {
                if (value.indexOf(currentEnvironmentId) === -1) {
                    continue
                }
                const valueJson = JSON.parse(value);
                if (Array.isArray(valueJson)) {
                    const environment = valueJson.filter(v => v.name === currentEnvironmentId)[0];
                    if (environment != null) {
                        return environment?.properties?.linkedEnvironmentMetadata?.instanceUrl;
                    }
                }
            } catch {
                // ignore
            }
        }
    }

    const modernPage = getModernPage(new URL(window.location.href));

    if (modernPage) {
        await openInWebApi(() => getModernWebApiUrl(modernPage));
    } else if (window.Xrm && window.Xrm.Page) {
        await openInWebApi(getLegacyWebApiUrl);
    } else if (/\/api\/data\/v\d+\.\d+\//.test(window.location.pathname)) {
        // the host check is for supporting on-prem, where we always want to resort to the postmessage based flow
        // we only need total reload on the workflows table when viewing a single record, because of the monaco editor
        if (window.location.hash === '#p' && /\/api\/data\/v\d+\.\d+\/workflows\(/.test(window.location.pathname) && window.location.host.endsWith(".dynamics.com")) {
            window.location.reload();
        } else {
            window.location.hash = 'p';
            window.postMessage({ action: 'prettifyWebApi' });
        }
    }
    else if (window.location.host.endsWith('.powerautomate.com') || window.location.host.endsWith('.powerapps.com')) {
        const hrefToCheck = location.href + '/'; // append a slash in case the url ends with '/flows' or '/cloudflows'

        if (hrefToCheck.indexOf('flows/') === -1 || hrefToCheck.indexOf('/environments/') === -1) {
            return;
        }

        const instanceUrl = getDataverseUrl();

        if (!instanceUrl) {
            console.warn(`PrettifyMyWebApi: Couldn't find Dataverse instanceUrl.`);
            return;
        }

        // it can be /cloudflows/ or /flows/ so just check for flows/
        const flowUniqueId = hrefToCheck.split('flows/').pop().split('?')[0].split('/')[0];

        if (flowUniqueId && guidPattern.test(flowUniqueId)) {
            let apiUrl;
            try {
                // absolute, whether or not the instanceUrl ends with a slash, e.g. https://contoso.crm4.dynamics.com/api/data/v9.2/
                apiUrl = new URL('/api/data/v9.2/', instanceUrl).href;
            } catch {
                console.warn(`PrettifyMyWebApi: The Dataverse instanceUrl is not a url: ${instanceUrl}`);
                return;
            }

            const url = apiUrl + 'workflows?$filter=resourceid eq ' + flowUniqueId + ' or workflowidunique eq ' + flowUniqueId + ' or workflowid eq ' + flowUniqueId + '#pf'
            window.postMessage({ action: 'openFlowInWebApi', url: url });
        } else {
            // for example, on make.powerapps it only works when viewing a flow from a solution
            console.warn(`PrettifyMyWebApi: Couldn't find Dataverse Flow Id.`);

            if (window.location.host.endsWith('make.powerapps.com') && hrefToCheck.indexOf('/solutions/') === -1) {
                alert(`Cannot find the Flow Id in the url. If you want to use this extension in make.powerapps.com, please open this Flow through a solution. Tip: if you use make.powerautomate.com, you should not run into this issue.`);
            }
        }
    }
})()