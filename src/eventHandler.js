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

            let recordId = Xrm.Page.data.entity.getId();
            recordId = recordId.replace("{", "");
            recordId = recordId.replace("}", "");

            const newLocation = window.location.origin + apiUrl + pluralName + "(" + recordId + ")";
            return newLocation;
        } catch (e) {
            alert("Error occurred: " + e.message);
        }
    }

    if (window.Xrm && window.Xrm.Page) {
        const newLocation = await getWebApiUrl();

        window.postMessage({ appId: "bhfdhnhbnamllpiaaapodfmoicgbbcmc", action: "openInWebApi", url: newLocation });
    } else if (/\/api\/data\/v[0-9][0-9]?.[0-9]\//.test(window.location.pathname)) {
        window.postMessage({ appId: "bhfdhnhbnamllpiaaapodfmoicgbbcmc", action: "prettifyWebApi" });
    }
})()
