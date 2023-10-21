(async function () {
    async function odataFetch(url) {
        const response = await fetch(url, { headers: { 'Prefer': 'odata.include-annotations="*"', 'Cache-Control': 'no-cache' } });

        return await response.json();
    }

    async function getWebApiUrl() {
        try {
            const versionArray = Xrm.Utility.getGlobalContext().getVersion().split('.');
            const version = versionArray[0] + '.' + versionArray[1];

            const entityLogicalName = Xrm.Page.data.entity.getEntityName();
            const apiUrl = "/api/data/v" + version + "/";

            const requestUrl = apiUrl + "EntityDefinitions?$select=EntitySetName&$filter=(LogicalName eq %27" + entityLogicalName + "%27)";

            const result = await odataFetch(requestUrl)
            const pluralName = result.value[0].EntitySetName;

            const recordId = Xrm.Page.data.entity.getId().replace("{", "").replace("}", "");

            const newLocation = window.location.origin + apiUrl + pluralName + "(" + recordId + ")";
            return newLocation;
        } catch (e) {
            alert("Error occurred: " + e.message);
        }
    }

    function getDataverseUrl() {
        const currentEnvironmentId = location.href.split("/environments/").pop().split("?")[0].split("/")[0];

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
        const newLocation = await getWebApiUrl() + '#p';
        window.postMessage({ action: "openInWebApi", url: newLocation });
    } else if (/\/api\/data\/v[0-9][0-9]?.[0-9]\//.test(window.location.pathname)) {
        window.location.hash = 'p';
        window.postMessage({ action: "prettifyWebApi" });
    }
    else if (window.location.host.endsWith(".powerautomate.com") || window.location.host.endsWith(".powerapps.com")) {
        const hrefToCheck = location.href + "/"; // append a slash in case the url ends with '/flows' or '/cloudflows'

        if (hrefToCheck.indexOf("flows/") === -1 || hrefToCheck.indexOf("/environments/") === -1) {
            return;
        }

        const instanceUrl = getDataverseUrl();

        if (!instanceUrl) {
            console.warn("PrettifyMyWebApi: Couldn't find Dataverse instanceUrl.");
            return;
        }

        // it can be /cloudflows/ or /flows/ so just check for flows/
        const flowUniqueId = hrefToCheck.split("flows/").pop().split("?")[0].split("/")[0];

        if (flowUniqueId && flowUniqueId.length === 36) {
            const url = instanceUrl + "api/data/v9.2/workflows?$filter=resourceid eq " + flowUniqueId + " or workflowidunique eq " + flowUniqueId + "#pf"
            window.postMessage({ action: "openFlowInWebApi", url: url });
        } else {
            // for example, on make.powerapps it only works when viewing a flow from a solution
            console.warn("PrettifyMyWebApi: Couldn't find Dataverse Flow Id.");

            if (window.location.host.endsWith("make.powerapps.com") && hrefToCheck.indexOf("/solutions/") === -1) {
                alert("Cannot find the Flow Id in the url. If you want to use this extension in make.powerapps.com, please open this Flow through a solution. Tip: if you use make.powerautomate.com, you should not run into this issue.");
            }
        }
    }
})()
