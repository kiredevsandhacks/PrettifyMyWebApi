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

			showNotification(e.message, `error`);
		}
	}


	function showNotification(message, type = 'info', time = 3500) {

		//this has been bothering me for a long time.
		if (message == "Cannot read properties of null (reading 'entity')") {
			message = "Please open a form to use PrettifyMyWebApi";
		}

		const SlideTime = 250;

		// Call the function to inject the element if not present yet!
		injectNotificationElement();

		// Construct class name
		var className = `pmwa-notification-${type}`;

		// Get the notification element
		var notification = document.getElementById("pmwa-notification");

		// Update the notification content
		notification.querySelector('span').innerHTML = message;

		// Update the notification class
		notification.className = ''; // Clear existing classes
		notification.classList.add(className); // Add new class

		// Show the notification
		notification.style.display = 'block';
		notification.style.opacity = 1;
		notification.style.transition = `opacity ${SlideTime}ms`;


		// Set a timer to hide the notification
		setTimeout(function () {
			notification.style.opacity = 0;
			notification.style.transition = `opacity ${SlideTime}ms`;
			setTimeout(function () {
				notification.style.display = 'none';
			}, SlideTime);
		}, time);
	}

	function injectNotificationElement() {

		// pmwa-notification already injected
		if (document.getElementById('pmwa-notificatio-notification')) { return; }

		// Create a new style element
		var styleElement = document.createElement('style');
		styleElement.textContent = `
        #pmwa-notification {
            font-family: SegoeUI, "Segoe UI";
            position: absolute;
            min-width: 200px;
            right: 30px;
            top: 80px;
            display: none;
             padding: 20px;
             font-size: 15px;
			 border-radius:5px;
        }

        .pmwa-notification-info {
	        background-color: #ccf6ff;
	        color: #056;
	        border: 1px solid #9ef;
        }

        .pmwa-notification-error {
	    	background-color: #ffd6e0;
	        color: #661429;
	        border:1px solid #ffadc2;
        }
    `;

		// Append the style element to the head
		document.head.appendChild(styleElement);

		// Create a new div element
		var notificationDiv = document.createElement('div');
		notificationDiv.id = 'pmwa-notification';

		// Create a new span element
		var spanElement = document.createElement('span');

		// Append the span to the div
		notificationDiv.appendChild(spanElement);

		// Append the notification div to the body
		document.body.appendChild(notificationDiv);
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
