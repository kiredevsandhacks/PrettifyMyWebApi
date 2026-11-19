(async function () {
    const guidPattern = /^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$/i;
    const logicalNamePattern = /^[a-z0-9_]+$/i;

    // true when the url stays on the protocol and host of this page (including the port, for on-prem)
    function isOnThisHost(url) {
        const resolved = new URL(url, window.location.href);
        return resolved.protocol === window.location.protocol && resolved.host === window.location.host;
    }

    async function odataFetch(url) {
        const response = await fetch(url, { headers: { 'Prefer': 'odata.include-annotations="*"', 'Cache-Control': 'no-cache' } });

        return await response.json();
    }

    async function getViewWebApiUrl(entityLogicalName, viewId, viewType) {
        // the view id can be wrapped in braces
        if (!guidPattern.test(viewId.replace(/^\{(.*)\}$/, '$1'))) {
            throw 'unknown view id: ' + viewId;
        }

        let queryParamName = '';

        if (viewType == 4230) {
            queryParamName = 'userQuery'
        } else if (viewType == 1039) {
            queryParamName = 'savedQuery';
        } else {
            throw 'unknown view type: ' + viewType;
        }

        const baseUrl = await getWebApiBaseUrl(entityLogicalName);

        return `${baseUrl}?${queryParamName}=${viewId}`;;
    }

    async function getSingleRowApiUrl() {
        const entityLogicalName = Xrm.Page.data.entity.getEntityName();

        const baseUrl = await getWebApiBaseUrl(entityLogicalName);

        const recordId = Xrm.Page.data.entity.getId().replace('{', '').replace('}', '');

        return `${baseUrl}(${recordId})`;
    }

    async function getWebApiBaseUrl(entityLogicalName) {
        if (!logicalNamePattern.test(entityLogicalName)) {
            throw 'unknown table: ' + entityLogicalName;
        }

        const versionArray = Xrm.Utility.getGlobalContext().getVersion().split('.');
        const version = versionArray[0] + '.' + versionArray[1];

        // on-prem, the path starts with the name of the organization, e.g. /contoso/main.aspx
        const pathSegments = window.location.pathname.split('/');
        const apiUrl = pathSegments.length <= 2 ? `/api/data/v${version}/` : `/${pathSegments[1]}/api/data/v${version}/`;

        if (!isOnThisHost(apiUrl)) {
            throw 'Could not determine the Web Api url of this environment from ' + window.location.pathname;
        }

        const requestUrl = apiUrl + 'EntityDefinitions?$select=EntitySetName&$filter=(LogicalName eq %27' + entityLogicalName + '%27)';

        const result = await odataFetch(requestUrl)
        const pluralName = result.value[0].EntitySetName;

        const baseUrl = window.location.origin + apiUrl + pluralName;

        return baseUrl;
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

    if (window.Xrm && window.Xrm.Page) {
        //no data and no utility => return
        if (!window.Xrm.Page.data && !window.Xrm.Utility) {
            alert(`Please open a form or view to use PrettifyMyWebApi`);
            return;
        }

        let urlToOpen = '';

        const urlObj = new URL(window.location.href);
        const viewId = urlObj.searchParams.get('viewid');
        const entityLogicalName = urlObj.searchParams.get('etn');
        const viewType = urlObj.searchParams.get('viewType');

        try {
            // check if on view
            if (viewId && entityLogicalName && viewType) {
                urlToOpen = await getViewWebApiUrl(entityLogicalName, viewId, viewType);
            } else if (window.Xrm.Page.data?.entity) {
                urlToOpen = await getSingleRowApiUrl();
            } else {
                alert(`Please open a form or view to use PrettifyMyWebApi`);
                return;
            }
        } catch (err) {
            alert(err);
            return;
        }

        urlToOpen += '#p'; // add the secret sauce
        window.postMessage({ action: 'openInWebApi', url: urlToOpen });
    } else if (/\/api\/data\/v[0-9][0-9]?.[0-9]\//.test(window.location.pathname)) {
        // the host check is for supporting on-prem, where we always want to resort to the postmessage based flow
        // we only need total reload on the workflows table when viewing a single record, because of the monaco editor
        if (window.location.hash === '#p' && /\/api\/data\/v[0-9][0-9]?.[0-9]\/workflows\(/.test(window.location.pathname) && window.location.host.endsWith(".dynamics.com")) {
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
            const url = instanceUrl + 'api/data/v9.2/workflows?$filter=resourceid eq ' + flowUniqueId + ' or workflowidunique eq ' + flowUniqueId + ' or workflowid eq ' + flowUniqueId + '#pf'
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