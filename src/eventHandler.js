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
			
			const apiUrl = window.location.pathname.split('/').length <= 2 ? `/api/data/v${version}/` : `/${window.location.pathname.split('/')[1]}/api/data/v${version}/`

			const requestUrl = apiUrl + 'EntityDefinitions?$select=EntitySetName&$filter=(LogicalName eq %27' + entityLogicalName + '%27)';

			const result = await odataFetch(requestUrl)
			const pluralName = result.value[0].EntitySetName;

			const recordId = Xrm.Page.data.entity.getId().replace('{', '').replace('}', '');

			const newLocation = window.location.origin + apiUrl + pluralName + '(' + recordId + ')';
			return newLocation;
		} catch (e) {
			alert('Error occurred: ' + e.message);
		}
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
	
	/**       
	Retrieve a record from the web api
	@param entityName The name of the entity without the trailing 's'
	@param id Guid of the record
	@param cols columns to retrieve
	@param keyAttribute keyAttribute defaults to entity + 'id', use this to override key attribute (ex activity)
	*/
	function Retrieve(entityName/*: string*/, id/*: string*/, cols/*?: Array<string>*/ = [], keyAttribute/*?: string*/ = ``) {
		if (!keyAttribute) {
			keyAttribute = entityName + "id";
		}

		if (entityName.substr(entityName.length - 1) == "y" && !entityName.endsWith('journey'))
		{
			entityName = entityName.substr(0, entityName.length - 1) + "ies";
		}

		else {

			entityName = entityName + "s";
		}

		id = id.replace(/[{}]/g, "").toLowerCase();

		var select = "$select=" + keyAttribute;

		if (cols) {
			select += ",";
			select += cols.join(',');
		}

		var req = new XMLHttpRequest();
		var clientURL = Xrm.Utility.getGlobalContext().getClientUrl();
		req.open("GET", encodeURI(clientURL + "/api/data/v9.2/" + entityName + "(" + id + ")?" + select), false);
		req.setRequestHeader("Accept", "application/json");
		req.setRequestHeader("Content-Type", "application/json; charset=utf-8");
		req.setRequestHeader("OData-MaxVersion", "4.0");
		req.setRequestHeader("OData-Version", "4.0");
		req.setRequestHeader("Prefer", 'odata.include-annotations="OData.Community.Display.V1.FormattedValue"');
		req.send(null);
		return JSON.parse(req.responseText);
	}

	function getPluralNameFromFetchXML(fetchXML) {
		//todo: fix this later
		const regex = /<entity\s+name="([^"]+)"/;
		const match = regex.exec(fetchXML);
		let entityName = match[1];
		if (entityName.substr(entityName.length - 1) == "y" && !entityName.endsWith('journey')) {
			entityName = entityName.substr(0, entityName.length - 1) + "ie";
		}

		const pluralName = `${entityName}s`
		return pluralName;
	}

	if (window.Xrm && window.Xrm.Page) {

		//no data and no utility => return
		if (!window.Xrm.Page.data && !window.Xrm.Utility) { 
			return;
		}

		//no data but utility ? check if on view
		if (!window.Xrm.Page.data && window.Xrm.Utility) {
			try {
				//Feature request:check if we are on a view
				// Get the current window URL
				const currentUrl = window.location.href;

				// Create a URL object from the current URL
				const urlObj = new URL(currentUrl);

				// Get the value of the 'viewid' parameter from the URL search parameters
				const viewId = urlObj.searchParams.get('viewid');

				//if on view, get the fetchxml
				if (viewId) {

					const currentview = Retrieve(`savedquery`, viewId, [`layoutjson`, 'layoutxml', `fetchxml`]);
					const fetchXml = currentview["fetchxml"]
					const encodedFetchXml = encodeURIComponent(fetchXml);
					const versionArray = Xrm.Utility.getGlobalContext().getVersion().split('.');
					const version = versionArray[0] + '.' + versionArray[1];
					const entityPluralName = getPluralNameFromFetchXML(fetchXml);
					const apiUrl = window.location.pathname.split('/').length <= 2 ?
						`/api/data/v${version}/${entityPluralName}?fetchXml=${encodedFetchXml}` :
						`/${window.location.pathname.split('/')[1]}/api/data/v${version}/${entityPluralName}?fetchXml=${encodedFetchXml}`
					const newLocation = Xrm.Utility.getGlobalContext().getClientUrl() + apiUrl + '#p';
					window.postMessage({ action: 'openInWebApi', url: newLocation });
					return;
				}

			} catch (e) {
				alert(`Please open a form or view to use PrettifyMyWebApi`);
				return;
			}
		}

		//in case we are not on form or on view
		if (!window.Xrm.Page.data || !window.Xrm.Page.data.entity) {

			alert(`Please open a form or view to use PrettifyMyWebApi`);
			return;
		}

		const newLocation = await getWebApiUrl() + '#p';
		window.postMessage({ action: 'openInWebApi', url: newLocation });
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

		if (flowUniqueId && flowUniqueId.length === 36) {
			const url = instanceUrl + 'api/data/v9.2/workflows?$filter=resourceid eq ' + flowUniqueId + ' or workflowidunique eq ' + flowUniqueId + '#pf'
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