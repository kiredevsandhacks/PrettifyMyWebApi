(async function () {
    const formattedValueType = '@OData.Community.Display.V1.FormattedValue';
    const navigationPropertyType = '@Microsoft.Dynamics.CRM.associatednavigationproperty';
    const lookupType = '@Microsoft.Dynamics.CRM.lookuplogicalname';

    const replacedQuote = '__~~__REPLACEDQUOTE__~~__';
    const replacedComma = '__~~__REPLACEDCOMMA__~~__';

    const clipBoardIcon = `<svg style='width:16px;position:absolute' viewBox='0 0 24 24'>
    <path fill='currentColor' d='M19,3H14.82C14.25,1.44 12.53,0.64 11,1.2C10.14,1.5 9.5,2.16 9.18,3H5A2,2 0 0,0 3,5V19A2,2 0 0,0 5,21H19A2,2 0 0,0 21,19V5A2,2 0 0,0 19,3M12,3A1,1 0 0,1 13,4A1,1 0 0,1 12,5A1,1 0 0,1 11,4A1,1 0 0,1 12,3M7,7H17V5H19V19H5V5H7V7M17,11H7V9H17V11M15,15H7V13H15V15Z' />
</svg>`.replaceAll(',', replacedComma); // need to 'escape' the commas because they cause issues with the JSON string cleanup code 

    let apiUrl = '';

    try {
        apiUrl = /\/api\/data\/v[0-9][0-9]?.[0-9]\//.exec(window.location.pathname)[0];
    } catch {
        alert('It seems you are not viewing a form or the dataverse odata web api. If you think this is an error, please contact the author of the extension and he will fix it asap.');
        return;
    }

    const retrievedPrimaryNamesAndKeys = {};
    const lazyLookupInitFunctions = {};
    const retrievedSearchableAttributes = {};
    let allPluralNames = null;

    async function odataFetch(url) {
        const response = await fetch(url, { headers: { 'Prefer': 'odata.include-annotations="*"', 'Cache-Control': 'no-cache' } });

        return await response.json();
    }

    async function retrievePluralName(logicalName) {
        const pluralNames = await retrieveAllPluralNames();

        const table = pluralNames.find(p => p.LogicalName === logicalName);

        return table.EntitySetName;
    }

    async function retrieveAllPluralNames() {
        if (allPluralNames != null) {
            return allPluralNames;
        }

        const requestUrl = apiUrl + "EntityDefinitions?$select=EntitySetName,LogicalName";

        const json = await odataFetch(requestUrl);

        allPluralNames = json.value;

        return allPluralNames;
    }

    // TODO: if lookup query is implemented, maybe refactor this into the retrieveAllPluralNames function? Let it just retrieve all needed metadata from the system in one api call
    async function retrievePrimaryNameAndKeyAndPluralName(logicalName) {
        if (retrievedPrimaryNamesAndKeys.hasOwnProperty(logicalName)) {
            return retrievedPrimaryNamesAndKeys[logicalName];
        }

        const requestUrl = apiUrl + "EntityDefinitions?$select=PrimaryNameAttribute,PrimaryIdAttribute,EntitySetName&$filter=(LogicalName eq '" + logicalName + "')";

        const json = await odataFetch(requestUrl);

        let primaryNameAndKey = {};
        primaryNameAndKey.name = json.value[0].PrimaryNameAttribute;
        primaryNameAndKey.key = json.value[0].PrimaryIdAttribute;
        primaryNameAndKey.plural = json.value[0].EntitySetName;

        retrievedPrimaryNamesAndKeys[logicalName] = primaryNameAndKey;

        return primaryNameAndKey;
    }

    async function retrieveLogicalNameFromPluralNameAsync(pluralName) {
        const requestUrl = apiUrl + "EntityDefinitions?$select=LogicalName,PrimaryIdAttribute&$filter=(EntitySetName eq '" + pluralName + "')";

        const json = await odataFetch(requestUrl);

        if (json.value.length === 0) {
            return {};
        }

        const logicalName = json.value[0].LogicalName;
        const primaryIdAttribute = json.value[0].PrimaryIdAttribute;

        return {
            logicalName: logicalName,
            primaryIdAttribute: primaryIdAttribute
        };
    }

    async function retrieveSearchableAttributes(logicalName) {
        if (retrievedSearchableAttributes.hasOwnProperty(logicalName)) {
            return retrievedSearchableAttributes[logicalName];
        }

        const requestUrl = apiUrl + "EntityDefinitions(LogicalName='" + logicalName + "')/Attributes?$filter=IsValidForAdvancedFind/Value eq true";

        const json = await odataFetch(requestUrl);

        retrievedSearchableAttributes[logicalName] = json.value;

        return json.value;
    }

    async function retrieveUpdateableAttributes(logicalName) {
        const requestUrl = apiUrl + "EntityDefinitions(LogicalName='" + logicalName + "')/Attributes?$filter=IsValidForUpdate eq true";

        const json = await odataFetch(requestUrl);

        return json.value;
    }

    async function retrieveOptionSetMetadata(logicalName, fieldname) {
        const requestUrl = apiUrl + "EntityDefinitions(LogicalName='" + logicalName + "')/Attributes/Microsoft.Dynamics.CRM.PicklistAttributeMetadata?$select=LogicalName&$expand=OptionSet,GlobalOptionSet";

        const json = await odataFetch(requestUrl);

        return json.value;
    }

    async function retrieveBooleanFieldMetadata(logicalName, fieldname) {
        const requestUrl = apiUrl + "EntityDefinitions(LogicalName='" + logicalName + "')/Attributes/Microsoft.Dynamics.CRM.BooleanAttributeMetadata?$select=LogicalName&$expand=OptionSet,GlobalOptionSet";

        const json = await odataFetch(requestUrl);

        return json.value;
    }


    async function doLookupInputQuery(pluralName, fieldToFilterOn, fieldToDisplay, primaryKeyfieldname, query) {
        query = query.replaceAll(`'`, `''`);
        const requestUrl = apiUrl + `${pluralName}?$top=10&$select=${fieldToDisplay},${primaryKeyfieldname}&$filter=contains(${fieldToFilterOn}, '${query}')`;

        const json = await odataFetch(requestUrl);

        return json.value;
    }

    const entityMap = {
        '&': '&amp;',
        '<': '&lt;',
        '>': '&gt;',
        '"': '&quot;',
        "'": '&#39;',
        '/': '&#x2F;',
        '`': '&#x60;',
        '=': '&#x3D;'
    };

    function escapeHtml(string) {
        return String(string).replace(/[&<>"'`=\/]/g, function (s) {
            return entityMap[s];
        });
    }

    function isNumber(n) {
        return typeof n == 'number' && !isNaN(n) && isFinite(n);
    }

    function determineType(value) {
        let cls = 'string';
        if (isNumber(value)) {
            cls = 'number';
        } else if (value === null) { // null check is explicit here (===) to prevent matching undefined (which renders as string)
            cls = 'null';
        } else if (typeof (value) === typeof (true)) {
            cls = 'boolean';
        }

        return cls;
    }

    function addcss(css) {
        const head = document.getElementsByTagName('head')[0];
        const s = document.createElement('style');
        s.setAttribute('type', 'text/css');
        if (s.styleSheet) {
            s.styleSheet.cssText = css;
        } else {
            s.appendChild(document.createTextNode(css));
        }
        head.appendChild(s);
    }

    function keyHasLookupAnnotation(key, jsonObj) {
        return !!(jsonObj[key + lookupType]);
    }

    function keyHasFormattedValueAnnotation(key, jsonObj) {
        return !!(jsonObj[key + formattedValueType]);
    }

    async function generateApiAnchorAsync(logicalName, guid) {
        const pluralName = await retrievePluralName(logicalName);
        const formattedGuid = guid.replace('{', '').replace('}', '');
        const newLocation = apiUrl + escapeHtml(pluralName) + '(' + escapeHtml(formattedGuid) + ')';

        return `<a target='_blank' href='${newLocation}'>${escapeHtml(formattedGuid)}</a>`;
    }

    function generateFormUrlAnchor(logicalName, guid) {
        const newLocation = '/main.aspx?etn=' + escapeHtml(logicalName) + '&id=' + escapeHtml(guid) + '&pagetype=entityrecord';

        return `<a target='_blank' href='${newLocation}'>Open in Form</a>`;
    }

    function generateWebApiAnchor(guid, pluralName) {
        const formattedGuid = guid.replace('{', '').replace('}', '');
        const newLocation = apiUrl + escapeHtml(pluralName) + '(' + escapeHtml(formattedGuid) + ')';

        return `<a target='_blank' href='${newLocation}'>Open in Web Api</a>`;
    }

    async function generatePreviewUrlAnchor(logicalName, guid) {
        const pluralName = await retrievePluralName(logicalName);
        const formattedGuid = guid.replace('{', '').replace('}', '');

        return `<a class='previewLink' data-pluralName='${escapeHtml(pluralName)}' data-guid='${escapeHtml(formattedGuid)}' href='javascript:'>Preview</a>`;
    }

    async function generateEditAnchor(logicalName, guid) {
        const pluralName = await retrievePluralName(logicalName);
        const formattedGuid = guid.replace('{', '').replace('}', '');

        return `
<a class='editLink' data-logicalName='${escapeHtml(logicalName)}' data-pluralName='${escapeHtml(pluralName)}' data-guid='${escapeHtml(formattedGuid)}' href='javascript:'>Edit this record</a>     
<div class='editMenuDiv' style='display: none;'>
    <div>    Bypass Custom Plugin Execution<input class='bypassPluginExecutionBox' type='checkbox' style='width:25px;'>
    </div><div>    Preview changes before committing save<input class='previewChangesBeforeSavingBox' type='checkbox' style='width:25px;' checked='true'>
    </div><div>    Impersonate another user<input class='impersonateAnotherUserCheckbox' type='checkbox' style='width:25px;'>
    </div><div class='impersonateDiv' style='display:none;'><div>      Base impersonation on this field: <select  class='impersonateAnotherUserSelect'><option value='systemuserid'>systemuserid</option><option value='azureactivedirectoryobjectid'>azureactivedirectoryobjectid</option></select>  <i><a href='https://learn.microsoft.com/en-us/power-apps/developer/data-platform/webapi/impersonate-another-user-web-api#how-to-impersonate-a-user' target='_blank'>What's this?</a></i>
    </div><div>      <span class='impersonationIdFieldLabel'>systemuserid:</span><input class='impersonateAnotherUserInput' placeholder='00000000-0000-0000-0000-000000000000'>  <span class='impersonateUserPreview'></span>
    </div></div><div><div id='previewChangesDiv'></div>    <a class='submitLink' style='display: none;' href='javascript:'>Save</a>
    </div>
</div>`.replaceAll('\n', '');
    }

    function createSpan(cls, value) {
        return `<span style='display: inline-flex;' class='${escapeHtml(cls)} field'>${escapeHtml(value)}<span class='copyButton'>` + clipBoardIcon + `</span></span>`;
    }

    function createSpanForLookup(cls, value) {
        return `<span style='display: inline-flex;' class='${escapeHtml(cls)} field lookupField'>${escapeHtml(value)}<span class='copyButton'>` + clipBoardIcon + `</span></span>`;
    }

    function createLinkSpan(cls, value) {
        // unsafe contents of 'value' have been escaped in a previous stage
        return `<span class='${escapeHtml(cls)}'>${value}</span>`;
    }

    function createFieldSpan(cls, value, fieldName) {
        return `<span style='display: inline-flex;' class='${escapeHtml(cls)} field'>${escapeHtml(value)}<div class='inputContainer containerNotEnabled' style='display: none;' data-fieldname='${escapeHtml(fieldName)}'></div><span class='copyButton'>` + clipBoardIcon + `</span></span>`;
    }

    function createLookupEditField(displayName, guid, fieldname, lookupTypeValue, navigationPropertyValue) {
        if (displayName == null) {
            displayName = '';
        }

        fieldname = fieldname.substring(1, fieldname.length - 1).substring(0, fieldname.length - 7);

        const formattedGuid = guid?.replace('{', '')?.replace('}', '');
        return `<div class='lookupEditLinks' style='display:none;' data-fieldname='${escapeHtml(fieldname)}'><span class='link'>   <a href='javascript:' class='searchDifferentRecord lookupEditLink' data-fieldname='${escapeHtml(fieldname)}'>Edit lookup</a></span><span class='link'>   <a href='javascript:' class='clearLookup lookupEditLink' data-fieldname='${escapeHtml(fieldname)}'>Clear lookup</a></span><span class='link'>   <a href='javascript:' class='cancelLookupEdit lookupEditLink' data-fieldname='${escapeHtml(fieldname)}' style='display:none;'>Undo changes</a></span></div><span class='lookupEdit' style='display: none;'><div class='inputContainer containerNotEnabled' data-name='${escapeHtml(displayName)}' data-id='${escapeHtml(formattedGuid)}' data-fieldname='${escapeHtml(fieldname)}' data-lookuptype='${escapeHtml(lookupTypeValue)}' data-navigationproperty='${escapeHtml(navigationPropertyValue)}'></div></span>`;
    }

    function createdFormattedValueSpan(cls, value, fieldName, formattedValue) {
        let insertedValue = '';

        // toString the value because it can be a number. The formattedValue is always a string
        if (value?.toString() !== formattedValue) {
            insertedValue = value + ' : ' + formattedValue;
        } else {
            insertedValue = value;
        }

        return `<span style='display: inline-flex;' class='${escapeHtml(cls)} field'>${escapeHtml(insertedValue)}<div class='inputContainer containerNotEnabled' style='display: none;' data-fieldname='${escapeHtml(fieldName)}'></div><span class='copyButton'>` + clipBoardIcon + `</span></span>`;
    }

    async function enrichObjectWithHtml(jsonObj, logicalName, pluralName, primaryIdAttribute, isSingleRecord, isNested, nestedLevel) {
        const recordId = jsonObj[primaryIdAttribute]; // we need to get this value before parsing or else it will contain html

        const ordered = orderProperties(jsonObj);

        for (let key in ordered) {
            let value = ordered[key];

            const cls = determineType(value);

            if (Array.isArray(value)) {
                ordered[key] = [];

                if (Object.values(value).every(v => typeof (v) === 'object')) {
                    for (let nestedKey in value) {
                        let nestedValue = value[nestedKey];

                        ordered[key][nestedKey] = await enrichObjectWithHtml(nestedValue, null, null, null, null, true, nestedLevel + 1);
                    }
                }
                else {
                    // if every value is not an object, it an an array of primitives. We want to render as an array with formatted values, without numbers prepended
                    for (let nestedKey in value) {
                        let nestedValue = value[nestedKey];

                        ordered[key].push(createSpan(cls, nestedValue));
                    }
                }

                continue;
            }

            if (typeof (value) === 'object' && value != null) {
                ordered[key] = await enrichObjectWithHtml(value, null, null, null, null, true, nestedLevel + 1);
                continue;
            }

            if (isAnnotation(key)) {
                continue;
            }

            if (value != null && value.replaceAll) {
                value = value.replaceAll(replacedQuote, ''); // to prevent malformed html and potential xss, disallow this string
                value = value.replaceAll('"', replacedQuote);
                value = value.replaceAll(',', replacedComma);
            }

            // this code is to fix the layout of lookups, 'manually' adding the spaces
            // hacky, but it works
            const increment = nestedLevel == 1 ? 0 : 3;
            const spaces = new Array(1 + increment + nestedLevel * 3).join(" ");

            if (keyHasLookupAnnotation(key, ordered)) {
                const formattedValueValue = ordered[key + formattedValueType];
                const navigationPropertyValue = ordered[key + navigationPropertyType];
                const lookupTypeValue = ordered[key + lookupType];

                const newApiUrl = await generateApiAnchorAsync(lookupTypeValue, value);
                const formUrl = generateFormUrlAnchor(lookupTypeValue, value);
                const previewUrl = await generatePreviewUrlAnchor(lookupTypeValue, value);

                let lookupFormatted = '';

                lookupFormatted += `<span class='lookupDisplay'>{<br>   ` + spaces +
                    createLinkSpan('link', newApiUrl) + ' - ' +
                    createLinkSpan('link', formUrl) + ' - ' +
                    createLinkSpan('link', previewUrl);
                lookupFormatted += '<br>   '
                lookupFormatted += spaces + createSpan(determineType(formattedValueValue), 'Name: ' + formattedValueValue);
                lookupFormatted += '<br>   '
                lookupFormatted += spaces + createSpan(determineType(lookupTypeValue), 'LogicalName: ' + lookupTypeValue);
                lookupFormatted += '<br>   '
                lookupFormatted += spaces + createSpan(determineType(navigationPropertyValue), 'NavigationProperty: ' + navigationPropertyValue);
                lookupFormatted += '<br>'
                lookupFormatted += spaces + '}';
                lookupFormatted += '</span>';
                lookupFormatted += `<span class='lookupEdit' style='display:none;' >`;
                lookupFormatted += createSpanForLookup('string', formattedValueValue);
                lookupFormatted += '</span>';
                lookupFormatted += createLookupEditField(formattedValueValue, value, key, lookupTypeValue, navigationPropertyValue);
                ordered[key] = lookupFormatted;

                delete ordered[key + formattedValueType];
                delete ordered[key + navigationPropertyType];
                delete ordered[key + lookupType];
            }
            else if (key.startsWith('_') && key.endsWith('_value')) {
                // we have a null lookup here
                ordered[key] = createSpanForLookup('null', value) + createLookupEditField(null, null, key, null, null);
            }
            else if (keyHasFormattedValueAnnotation(key, ordered)) {
                ordered[key] = createdFormattedValueSpan(cls, value, key, ordered[key + formattedValueType]);
                delete ordered[key + formattedValueType];
            } else {
                if (key === primaryIdAttribute) {
                    ordered[key] = '<b>' + createSpan('primarykey', value) + '</b>';
                } else {
                    ordered[key] = createFieldSpan(cls, value, key);
                }
            }
        }

        const newObj = {};
        if (!isNested) {
            if (logicalName != null && logicalName !== '' && recordId != null && recordId !== '') {
                newObj['Form Link'] = createLinkSpan('link', generateFormUrlAnchor(logicalName, recordId));

                if (isSingleRecord) {
                    newObj['Edit this record'] = createLinkSpan('link', await generateEditAnchor(logicalName, recordId));
                } else {
                    newObj['Web Api Link'] = createLinkSpan('link', generateWebApiAnchor(recordId, pluralName));
                }
            } else if (logicalName != null && logicalName !== '' && (recordId == null || recordId === '')) {
                newObj['Form Link'] = 'Could not generate link';
                newObj['Web Api Link'] = 'Could not generate link';
            }
        }

        const combinedJsonObj = Object.assign(newObj, ordered);
        return combinedJsonObj;
    }

    function orderProperties(jsonObj) {
        return Object.keys(jsonObj).sort(
            (obj1, obj2) => {
                let obj1Underscore = obj1.startsWith('_');
                let obj2Underscore = obj2.startsWith('_');
                if (obj1Underscore && !obj2Underscore) {
                    return 1;
                } else if (!obj1Underscore && obj2Underscore) {
                    return -1;
                }
                else
                    return obj1 > obj2 ? 1 : -1;

            }).reduce(
                (obj, key) => {
                    obj[key] = jsonObj[key];
                    return obj;
                },
                {}
            );
    }

    function setCopyToClipboardHandlers() {
        Array.from(document.querySelectorAll('.copyButton')).forEach((el) => el.onclick = (element) => {
            navigator.clipboard.writeText(el.parentElement.innerText).then(() => {
                console.log('Content copied to clipboard');
            }, () => {
                alert('Failed to copy');
            });
        }
        );
    }

    function setPreviewLinkClickHandlers() {
        const previewLinks = document.getElementsByClassName('previewLink');

        for (let previewLink of previewLinks) {
            const pluralName = previewLink.attributes['data-pluralName'].value;
            const newLocation = pluralName + '(' + previewLink.attributes['data-guid'].value + ')';

            previewLink.onclick = function () {
                previewRecord(pluralName, newLocation);
            }
        }
    }

    function setEditLinkClickHandlers() {
        const editLinks = document.getElementsByClassName('editLink');

        // TODO refactor .attributes to .dataset
        for (let editLink of editLinks) {
            const logicalName = editLink.attributes['data-logicalName'].value;
            const pluralName = editLink.attributes['data-pluralName'].value;
            const id = editLink.attributes['data-guid'].value;

            editLink.onclick = async function () {
                await editRecord(logicalName, pluralName, id);
            }
        }
    }

    function setImpersonateUserHandlers() {
        const impersonateAnotherUserCheckbox = document.getElementsByClassName('impersonateAnotherUserCheckbox')[0];
        const impersonateAnotherUserSelect = document.getElementsByClassName('impersonateAnotherUserSelect')[0];
        const impersonationIdFieldLabel = document.getElementsByClassName('impersonationIdFieldLabel')[0];
        const impersonateDiv = document.getElementsByClassName('impersonateDiv')[0];
        const impersonateAnotherUserInput = document.getElementsByClassName('impersonateAnotherUserInput')[0];
        const impersonateUserPreview = document.getElementsByClassName('impersonateUserPreview')[0];

        impersonateAnotherUserSelect.onchange = async () => {
            impersonationIdFieldLabel.innerText = impersonateAnotherUserSelect.value + ':';
            await handleUserPreview();
        }

        impersonateAnotherUserCheckbox.onclick = () => {
            if (!!impersonateAnotherUserCheckbox.checked) {
                impersonateDiv.style.display = 'inline';
            } else {
                impersonateDiv.style.display = 'none'
            }
        }

        impersonateAnotherUserInput.oninput = async () => {
            await handleUserPreview();
        };

        async function handleUserPreview() {
            if (!impersonateAnotherUserInput.value) {
                impersonateUserPreview.innerText = '';
                return;
            }
            const retrievedSystemUser = await odataFetch(apiUrl + `systemusers?$filter=${impersonateAnotherUserSelect.value} eq '${impersonateAnotherUserInput.value}'&$select=fullname`);
            if (retrievedSystemUser.error) {
                impersonateUserPreview.innerText = retrievedSystemUser.error.message;
            } else if (retrievedSystemUser.value.length == 0) {
                impersonateUserPreview.innerText = 'user not found';
            } else if (retrievedSystemUser.value.length == 1) {
                impersonateUserPreview.innerText = retrievedSystemUser.value[0].fullname;
            } else {
                impersonateUserPreview.innerText = 'Something went wrong with retrieving the systemuser.';
            }
        }
    }

    function setLookupQueryByIdHandlers(input, lookupQueryResultPreview, selectTable) {
        input.oninput = async () => await handlePreview();
        selectTable.onchange = async () => await handlePreview();

        async function handlePreview() {
            // first fire this method so that the proper logicalname and pluralname is added to the input dataset
            await handleTableSelect(selectTable, input);

            if (!input.value) {
                lookupQueryResultPreview.innerText = '';
                return;
            }

            if (!selectTable.value) {
                alert('selectTable does not contain a value. This should not happen.');
                return;
            }

            const { plural, key, name } = await retrievePrimaryNameAndKeyAndPluralName(selectTable.value);

            const retrievedRecord = await odataFetch(apiUrl + `${plural}(${input.value})?$select=${name}`);

            if (retrievedRecord.error) {
                lookupQueryResultPreview.innerText = retrievedRecord.error.message;
            } else {
                lookupQueryResultPreview.innerText = retrievedRecord[name];
            }

            //  else if (retrievedRecord.value.length == 1) {
            //     lookupQueryResultPreview.innerText = retrievedRecord.value[0][name];
            // } else {
            //     lookupQueryResultPreview.innerText = 'Something went wrong with retrieving the systemuser.';
            // }
        }
    }

    function setLookupEditHandlers() {
        const clearLookupLinks = document.getElementsByClassName('clearLookup');

        for (let link of clearLookupLinks) {
            const logicalName = link.dataset.fieldname;
            link.onclick = () => {
                const enabledInputFields = document.getElementsByClassName('enabledInputField');
                const lookup = [...enabledInputFields].find(f => f.dataset.fieldname === logicalName);
                lookup.value = null;
                lookup.dataset.id = null;
                lookup.dataset.editmode = 'makenull';

                const makeNullDiv = [...document.getElementsByClassName('makeLookupNullDiv')].find(f => f.dataset.fieldname === logicalName);
                makeNullDiv.style.display = 'block';

                const lookupEditMenuDivs = document.getElementsByClassName('lookupEditMenuDiv');
                const lookupEditMenuDiv = [...lookupEditMenuDivs].find(f => f.dataset.fieldname === logicalName);

                lookupEditMenuDiv.style.display = 'none';

                const cancelLookupEditDivs = document.getElementsByClassName('cancelLookupEdit');
                const cancelLookupEditDiv = [...cancelLookupEditDivs].find(f => f.dataset.fieldname === logicalName);
                cancelLookupEditDiv.style.display = 'unset';
            };
        }

        const searchDifferentRecordLinks = document.getElementsByClassName('searchDifferentRecord');

        for (let link of searchDifferentRecordLinks) {
            const logicalName = link.dataset.fieldname;
            link.onclick = async () => {
                const enabledInputFields = document.getElementsByClassName('enabledInputField');
                const lookup = [...enabledInputFields].find(f => f.dataset.fieldname === logicalName);
                lookup.dataset.editmode = 'update';

                const lookupEditMenuDivs = document.getElementsByClassName('lookupEditMenuDiv');
                const lookupEditMenuDiv = [...lookupEditMenuDivs].find(f => f.dataset.fieldname === logicalName);

                lookupEditMenuDiv.style.display = 'unset';

                const makeNullDiv = [...document.getElementsByClassName('makeLookupNullDiv')].find(f => f.dataset.fieldname === logicalName);
                makeNullDiv.style.display = 'none';

                const cancelLookupEditDivs = document.getElementsByClassName('cancelLookupEdit');
                const cancelLookupEditDiv = [...cancelLookupEditDivs].find(f => f.dataset.fieldname === logicalName);
                cancelLookupEditDiv.style.display = 'unset';

                await lazyLookupInitFunctions[logicalName]?.call();
            };
        }

        const cancelLookupEditLinks = document.getElementsByClassName('cancelLookupEdit');

        for (let link of cancelLookupEditLinks) {
            const logicalName = link.dataset.fieldname;
            link.onclick = () => {
                const enabledInputFields = document.getElementsByClassName('enabledInputField');
                const lookup = [...enabledInputFields].find(f => f.dataset.fieldname === logicalName);
                lookup.value = lookup.dataset.originalname;
                lookup.dataset.id = lookup.dataset.originalid;
                lookup.dataset.editmode = null;

                link.style.display = 'none';

                const makeNullDiv = [...document.getElementsByClassName('makeLookupNullDiv')].find(f => f.dataset.fieldname === logicalName);
                makeNullDiv.style.display = 'none';

                const lookupEditMenuDivs = document.getElementsByClassName('lookupEditMenuDiv');
                const lookupEditMenuDiv = [...lookupEditMenuDivs].find(f => f.dataset.fieldname === logicalName);

                lookupEditMenuDiv.style.display = 'none';
            };
        }
    }

    function createInput(container, multiLine, datatype) {
        const value = window.originalResponseCopy[container.dataset.fieldname];

        let input;

        if (!multiLine) {
            input = document.createElement('input');
        } else {
            input = document.createElement('textarea');
        }

        input.value = value;

        setInputMetadata(input, container, datatype);
    }

    function createOptionSetValueInput(container, optionSet) {
        const value = window.originalResponseCopy[container.dataset.fieldname];

        const select = document.createElement('select');

        let selectHtml = "<option value='null'>null</option>"; // empty option for clearing it

        let cachedValue;

        optionSet.forEach(function (option) {
            const formattedOption = option.Value + ' : ' + option.Label.UserLocalizedLabel.Label;
            if (value === option.Value) {
                cachedValue = formattedOption;
            }
            // TODO: refactor the value attribute to contain the pure values, true/false/null
            selectHtml += `<option value='${escapeHtml(formattedOption)}'>${escapeHtml(formattedOption)}</option>`;
        });

        select.innerHTML = selectHtml;

        if (cachedValue) {
            select.value = cachedValue;
        }

        setInputMetadata(select, container, 'option');
    }

    function createBooleanInput(container, falseOption, trueOption) {
        const value = window.originalResponseCopy[container.dataset.fieldname];

        const select = document.createElement('select');

        let selectHtml = "<option value='null'>null</option>"; // empty option for clearing it

        const falseFormatted = 'false : ' + falseOption.Label.UserLocalizedLabel.Label;
        const trueFormatted = 'true : ' + trueOption.Label.UserLocalizedLabel.Label;
        // TODO: refactor the value attribute to contain the pure values, true/false/null
        selectHtml += `<option value='${escapeHtml(falseFormatted)}'>${escapeHtml(falseFormatted)}</option>`;
        selectHtml += `<option value='${escapeHtml(trueFormatted)}'>${escapeHtml(trueFormatted)}</option>`;
        select.innerHTML = selectHtml;

        if (value == null) {
            select.value = 'null';
        } else if (value === true) {
            select.value = trueFormatted;
        } else if (value === false) {
            select.value = falseFormatted;
        }

        setInputMetadata(select, container, 'bool');
    }

    function createLookupInput(container, targets) {
        let id = container.dataset.id;
        if (id == 'undefined' || id == undefined || id == null) {
            id = 'null';
        }

        const fieldName = container.dataset.fieldname;
        const name = container.dataset.name;
        const lookupType = container.dataset.lookuptype;
        const navigationProperty = container.dataset.navigationproperty;

        const input = document.createElement('input');
        input.placeholder = '00000000-0000-0000-0000-000000000000';

        input.dataset.id = id;
        input.dataset.originalid = id;
        input.dataset.logicalname = lookupType;
        input.dataset.originallogicalname = lookupType;

        const selectTable = document.createElement('select');
        selectTable.style.height = '22px';
        selectTable.classList.add('lookupSelectTable');
        selectTable.dataset.fieldname = fieldName;

        for (let target of targets) {
            let option = document.createElement('option');
            option.value = target;
            option.innerText = target;
            selectTable.appendChild(option);
        }

        input.dataset.ismultipletarget = false;
        if (targets.length < 2) {
            selectTable.setAttribute('disabled', 'disabled');
            input.dataset.ismultipletarget = 'false';
        } else {
            input.dataset.ismultipletarget = 'true';
        }

        const editMenuDiv = document.createElement('div');
        editMenuDiv.dataset.fieldname = fieldName;
        editMenuDiv.style.display = 'none';
        editMenuDiv.classList.add('lookupEditMenuDiv');

        const selectTableDiv = document.createElement('div');
        selectTableDiv.append('      table: ');
        selectTableDiv.append(selectTable);

        editMenuDiv.append(selectTableDiv);

        const makeNullDiv = document.createElement('div');
        makeNullDiv.dataset.fieldname = fieldName;
        makeNullDiv.style.display = 'none';
        makeNullDiv.classList.add('makeLookupNullDiv');

        makeNullDiv.innerHTML = '      <b>Lookup will be cleared (will be set to null)</b>'

        container.parentElement.appendChild(makeNullDiv);

        const lookupEditLinks = document.getElementsByClassName('lookupEditLinks');
        const lookupEditLinkDiv = [...lookupEditLinks].find(l => l.dataset.fieldname === fieldName);
        lookupEditLinkDiv.style.display = 'unset';

        setInputMetadataForLookup(input, container, editMenuDiv);

        const lookupQueryResultPreview = document.createElement('span');
        lookupQueryResultPreview.style.margin = '0 0 0 10px';
        editMenuDiv.appendChild(lookupQueryResultPreview);


        let targetToCache = lookupType;
        if (targetToCache === 'null' || targetToCache == null || targetToCache == undefined) {
            targetToCache = targets[0];
        }

        selectTable.value = targetToCache;

        initLookupMetadata(targetToCache, input);

        // fire the handle table select method so that we ensure that we have the proper dataset values for logicalname/pluralname on the input
        handleTableSelect(selectTable, input);

        setLookupQueryByIdHandlers(input, lookupQueryResultPreview, selectTable);
        return;
        // there is logic here for querying records. 
        // Not enabled for now as it's complicated and very hard to give a great user experience

        // placeholder after scrapping the lookup function. Needs to be placed somewhere
        const queryInput = document.createElement('input');

        const queryDiv = document.createElement('div');
        queryDiv.append('      query: ');
        queryDiv.append(queryInput);
        editMenuDiv.append(queryDiv);

        const selectFilterField = document.createElement('select');
        selectFilterField.dataset.fieldname = fieldName;
        selectFilterField.style.height = '22px';

        const resultsSelect = document.createElement('select');
        const resultsDiv = document.createElement('div');
        resultsDiv.append('      query results:');
        resultsDiv.append(resultsSelect);

        resultsSelect.dataset.id = id;
        resultsSelect.dataset.originalid = id;
        resultsSelect.dataset.originalname = name;
        resultsSelect.dataset.originalNavigationProperty = navigationProperty;
        resultsSelect.dataset.originalLookupType = lookupType;

        const selectFilterFieldDefault = document.createElement('option');
        selectFilterFieldDefault.value = '_primary';
        selectFilterFieldDefault.innerText = 'primary column';
        selectFilterField.appendChild(selectFilterFieldDefault);

        editMenuDiv.append(resultsDiv);

        const selectFilterFieldDiv = document.createElement('div');
        selectFilterFieldDiv.append('      field to query: ');
        selectFilterFieldDiv.append(selectFilterField);
        editMenuDiv.append(selectFilterFieldDiv);

        lazyLookupInitFunctions[fieldName] = async () => initLookupMetadata(targetToCache);

        setLookupInputQueryHandlers(queryInput, selectTable, selectFilterField, resultsSelect);
    }

    async function handleTableSelect(selectTable, input) {
        const logicalName = selectTable.value;
        const pluralName = await retrievePluralName(logicalName);
        input.dataset.pluralname = pluralName;
        input.dataset.logicalname = logicalName;
    }

    async function handleTableSelectv2(selectTable, selectFilterField) {
        const logicalName = selectTable.value;
        // retrieve the name etc. in advance because we will need it anyway
        const { name, key, plural } = await retrievePrimaryNameAndKeyAndPluralName(logicalName);
        const attributes = await retrieveSearchableAttributes(logicalName);

        selectFilterField.innerHTML = ''; // reset

        const selectFilterFieldDefault = document.createElement('option');
        selectFilterFieldDefault.value = '_primary';
        selectFilterFieldDefault.innerText = 'primary column (' + name + ')';
        selectFilterField.appendChild(selectFilterFieldDefault);

        const compare = (a, b) => {
            if (a.LogicalName < b.LogicalName) {
                return -1;
            }
            if (a.LogicalName > b.LogicalName) {
                return 1;
            }
            return 0;
        }

        const ordered = attributes.sort(compare);

        for (let attribute of ordered) {
            let selectFilterFieldOption = document.createElement('option');
            selectFilterFieldOption.value = attribute.LogicalName;
            selectFilterFieldOption.innerText = attribute.LogicalName;
            selectFilterField.appendChild(selectFilterFieldOption);
        }
    }

    async function initLookupMetadata(logicalName, input) {
        const pluralName = await retrievePluralName(logicalName);
        input.dataset.originalpluralname = pluralName;
        input.dataset.pluralname = pluralName;

        // these lines are turned off while lookup query is not enabled
        // await retrieveSearchableAttributes(logicalName);
        // await retrievePrimaryNameAndKeyAndPluralName(logicalName);
    }

    async function setLookupInputQueryHandlers(input, selectTable, selectFilterField, resultsSelect) {
        input.oninput = async () => {
            if (selectTable.value == null) {
                alert('Table to query is null. Cannot continue.')
                return;
            }
            if (selectFilterField.value == null) {
                alert('Field to query is null. Cannot continue.')
                return;
            }

            if (input.value == null || input.value == '') {
                resultsSelect.innerHTML = ''; // wipe it and stop
                return;
            }

            // the values retrieved by retrievePrimaryNameAndKeyAndPluralName are cached and the api call will already have been done at this point
            // unless the user selected a different table, in that case the api call will be done one time for the new table if not already cached
            const primaryNameAndKeyAndPlural = await retrievePrimaryNameAndKeyAndPluralName(selectTable.value);

            const primaryfieldname = primaryNameAndKeyAndPlural.name;
            const primaryKeyName = primaryNameAndKeyAndPlural.key;
            const pluralName = primaryNameAndKeyAndPlural.plural;

            let fieldToQuery = selectFilterField.value;

            if (fieldToQuery === '_primary') {
                fieldToQuery = primaryfieldname;
            }

            const results = await doLookupInputQuery(pluralName, fieldToQuery, primaryfieldname, primaryKeyName, input.value);

            resultsSelect.innerHTML = "<option value='null'></option>"; // empty option so that we don't select by default

            for (let result of results) {
                let option = document.createElement('option');
                option.value = `/${pluralName}/(${result[primaryKeyName]})`;
                option.innerText = result[primaryfieldname];
                resultsSelect.appendChild(option);
            }
        };
    }

    function setInputMetadataForLookup(input, container, editMenuDiv) {
        input.classList.add('enabledInputField');
        input.dataset.fieldname = container.dataset.fieldname;
        input.dataset.datatype = 'lookup';

        editMenuDiv.append('      record id:');
        editMenuDiv.appendChild(input);

        container.parentElement.append(editMenuDiv);

        container.style.display = null;
        container.classList.remove('containerNotEnabled');
        container.classList.add('containerEnabled');

        container.style.display = 'none';
    }

    function setInputMetadata(input, container, datatype) {
        input.classList.add('enabledInputField');
        input.dataset.fieldname = container.dataset.fieldname;
        input.dataset.datatype = datatype;

        container.parentElement.append(input);

        if (datatype === 'memo') {
            container.parentElement.style.display = null;
        }

        container.style.display = null;
        container.classList.remove('containerNotEnabled');
        container.classList.add('containerEnabled');
    }

    async function editRecord(logicalName, pluralName, id) {
        const editLink = document.getElementsByClassName('editLink')[0];
        editLink.style.display = 'none';

        const attributesMetadata = await retrieveUpdateableAttributes(logicalName);

        const optionSetMetadata = await retrieveOptionSetMetadata(logicalName);
        const booleanMetadata = await retrieveBooleanFieldMetadata(logicalName);

        const inputContainers = document.getElementsByClassName('mainPanel')[0].getElementsByClassName('inputContainer');
        const inputContainersArray = [...inputContainers];

        for (let attribute of attributesMetadata) {
            let container = inputContainersArray.find(c => c.dataset.fieldname === attribute.LogicalName);

            if (container == null) {
                continue;
            }

            const attributeType = attribute.AttributeType;
            if (attributeType === 'String') {
                createInput(container, false, 'string');
            } else if (attributeType === 'Memo') {
                createInput(container, true, 'memo');
            } else if (attributeType === 'Picklist') {
                const fieldOptionSetMetadata = optionSetMetadata.find(osv => osv.LogicalName === attribute.LogicalName);
                if (fieldOptionSetMetadata) {
                    const fieldOptionset = fieldOptionSetMetadata.GlobalOptionSet || fieldOptionSetMetadata.OptionSet;
                    createOptionSetValueInput(container, fieldOptionset.Options)
                }
            } else if (attributeType === 'Integer') {
                createInput(container, false, 'int');
            } else if (attributeType === 'Decimal') {
                createInput(container, false, 'decimal');
            } else if (attributeType === 'Money') {
                createInput(container, false, 'money');
            } else if (attributeType === 'Double') {
                createInput(container, false, 'float');
            } else if (attributeType === 'Boolean') {
                const fieldOptionSetMetadata = booleanMetadata.find(osv => osv.LogicalName === attribute.LogicalName);
                if (fieldOptionSetMetadata) {
                    const fieldOptionset = fieldOptionSetMetadata.GlobalOptionSet || fieldOptionSetMetadata.OptionSet;
                    createBooleanInput(container, fieldOptionset.FalseOption, fieldOptionset.TrueOption)
                }
            } else if (attributeType === 'Lookup') {
                const targets = attribute.Targets;
                
                // not all lookups are targeted to a table apparently, skip if there are no targets
                if (targets.length > 0) {
                    createLookupInput(container, targets);
                }
            } else if (attributeType === 'DateTime') {
                // todo
            } else if (attributeType === 'Uniqueidentifier') {
                // can't change this
            } else if (attributeType === 'State') {
                // difficult to implement
            } else if (attributeType === 'Status') {
                // difficult to implement
            } else if (attributeType === 'Virtual') {
                // difficult to implement
            }
        }

        const notEnabledContainers = document.getElementsByClassName('mainPanel')[0].getElementsByClassName('containerNotEnabled');

        const length = notEnabledContainers.length;
        for (let i = 0; i < length; i++) {
            notEnabledContainers[0].remove();
        }

        const editMenuDiv = document.getElementsByClassName('mainPanel')[0].getElementsByClassName('editMenuDiv')[0];
        editMenuDiv.style.display = 'inline';

        const submitLink = document.getElementsByClassName('mainPanel')[0].getElementsByClassName('submitLink')[0];
        submitLink.style.display = null;
        submitLink.onclick = async function () {
            await submitEdit(pluralName, id);
        }

        document.querySelectorAll('.field').forEach((el) => {
            // remove all hover handlers as they mess up the formatting and are not wanted in the editing context
            el.classList.remove('field');

            // set to fixed heigth for cleaner looking page
            el.style.height = '20px';

            if (el.classList.contains('lookupField')) {
                el.style.margin = '1px 0 -2px 0';
            }
        });

        document.querySelectorAll('.lookupEdit').forEach((el) => {
            el.style.display = 'unset';
        });

        document.querySelectorAll('.lookupDisplay').forEach((el) => {
            el.style.display = 'none';
        });
    }

    async function submitEdit(pluralName, id) {
        const previewChangesBeforeSaving = document.getElementsByClassName('mainPanel')[0].getElementsByClassName('previewChangesBeforeSavingBox')[0].checked;

        const changedFields = {};
        const enabledFields = document.getElementsByClassName('mainPanel')[0].getElementsByClassName('enabledInputField');
        for (let input of enabledFields) {
            let originalValue = window.originalResponseCopy[input.dataset.fieldname];
            const dataType = input.dataset.datatype;
            const inputValue = input.value;
            const fieldName = input.dataset.fieldname;

            let value = '';
            if (dataType === 'string' || dataType === 'memo') {
                if (inputValue === '') {
                    value = null;
                } else {
                    value = inputValue;
                }
            } else if (dataType === 'option') {
                if (!inputValue) {
                    // the select needs to contain a value always, if not, an error happened
                    alert('there was an error parsing the field ' + fieldName);
                    return;
                }
                if (inputValue === 'null') {
                    value = null;
                } else {
                    value = parseInt(inputValue.split(':')[0].replace(' ', ''));
                }
            } else if (dataType === 'int') {
                if (inputValue == null || inputValue === undefined || inputValue === '') {
                    value = null;
                } else {
                    value = parseInt(inputValue);
                    if (isNaN(value)) {
                        alert(fieldName + ' is a whole number. The value ' + inputValue + ' is not compatible.');
                        return;
                    }

                    if (/^-?\d+$/.test(inputValue) === false) {
                        alert(fieldName + ' is a whole number. The value ' + inputValue + ' is not compatible.');
                        return;
                    }
                }
            } else if (dataType === 'decimal' || dataType == 'money' || dataType == 'float') {
                if (inputValue == null || inputValue === undefined || inputValue === '') {
                    value = null;
                } else {
                    if (inputValue.includes(',')) {
                        alert(`${fieldName} is a column of type '${dataType}' and contains a comma (,). Use a dot (.) as the separator.`);
                        return;
                    }

                    value = parseFloat(inputValue);
                    if (isNaN(value)) {
                        alert(`${fieldName} is a column of type '${dataType}'. The value ${inputValue} is not compatible.`);
                        return;
                    }

                    if (/^-?[0-9]\d*(\.\d+)?$/.test(inputValue) === false) {
                        alert(`${fieldName} is a column of type '${dataType}'. The value ${inputValue} is not compatible.`);
                        return;
                    }
                }
            } else if (dataType === 'bool') {
                if (!inputValue) {
                    // the select needs to contain a value always, if not, an error happened
                    alert('there was an error parsing the field ' + fieldName);
                    return;
                }
                if (inputValue === 'null') {
                    value = null;
                } else {
                    let rawValue = inputValue.split(':')[0].replace(' ', '');
                    if (rawValue === 'true') {
                        value = true;
                    } else if (rawValue === 'false') {
                        value = false;
                    } else {
                        alert('there was an error parsing the field ' + fieldName);
                        return;
                    }
                }
            }
            else if (dataType === 'lookup') {
                // override the original value by taking the dataset values
                if (input.dataset.originalid !== 'null') {
                    originalValue = `/${input.dataset.originalpluralname}(${input.dataset.originalid})`;
                } else {
                    originalValue = null;
                }

                const tablesSelects = document.getElementsByClassName('lookupSelectTable');
                const tableSelect = [...tablesSelects].find(f => f.dataset.fieldname === fieldName);

                const tableSelectValue = tableSelect.value;

                if (tableSelectValue == null || tableSelectValue === '') {
                    alert('Error for lookup field: ' + fieldName + '. No table was selected. This should not be possible.');
                    return;
                }

                const pluralName = await retrievePluralName(tableSelectValue)

                value = `/${pluralName}(${inputValue})`;

                if (input.dataset.editmode === 'update') {
                    if (inputValue == null || inputValue === '') {
                        alert('Error for lookup field: ' + fieldName + '. The field was marked for update but it is empty. If you do not want to edit this field, hit "undo changes".');
                        return;
                    }

                    if (inputValue.length !== 36) {
                        alert('Error for lookup field: ' + fieldName + '. The value ' + value + ' is not a valid guid.');
                        return;
                    }
                } else if (input.dataset.editmode === 'makenull') {
                    value = null;
                } else {
                    // skip if not action is required
                    continue;
                }
            }


            if (value !== originalValue && !(value === '' && originalValue == null)) {
                if (dataType === 'memo') {
                    if (originalValue?.replaceAll('\r\n', '\n') !== value) {
                        changedFields[fieldName] = value;
                    }
                } else if (dataType == 'lookup') {
                    const isMultipleTarget = input.dataset.ismultipletarget;
                    let lookupFieldName = '';

                    // isMultipleTarget is set as a pure boolean true but will be reduced to a string because it is an attribute
                    if (isMultipleTarget === 'false') {
                        lookupFieldName = `${fieldName}@odata.bind`;
                    } else if (isMultipleTarget === 'true') {
                        lookupFieldName = `${fieldName}_${input.dataset.logicalname}@odata.bind`;
                    } else {
                        alert('Invalid value for isMultipleTarget.');
                        return;
                    }

                    if (!!previewChangesBeforeSaving) {
                        changedFields[lookupFieldName + '____lookupOverride'] = true;
                        changedFields[lookupFieldName + '____lookupOverrideOriginalValue'] = originalValue;
                    }

                    changedFields[lookupFieldName] = value;
                }
                else {
                    changedFields[fieldName] = value;
                }
            }
        }

        const impersonateAnotherUser = document.getElementsByClassName('impersonateAnotherUserCheckbox')[0].checked;
        const impersonateAnotherUserField = document.getElementsByClassName('impersonateAnotherUserSelect')[0].value;
        const impersonateAnotherUserInput = document.getElementsByClassName('impersonateAnotherUserInput')[0].value;

        const impersonateHeader = {};
        if (!!impersonateAnotherUser) {
            if (impersonateAnotherUserInput == null || impersonateAnotherUserInput === '') {
                alert('User impersonation was checked, but ' + impersonateAnotherUserField + ' is empty');
                return;
            }

            if (impersonateAnotherUserInput?.length !== 36) {
                alert('User impersonation input error: ' + impersonateAnotherUserInput + ' is not a valid guid.');
                return;
            }

            if (impersonateAnotherUserField == 'systemuserid') {
                impersonateHeader['MSCRMCallerID'] = impersonateAnotherUserInput;
            } else if (impersonateAnotherUserField == 'azureactivedirectoryobjectid') {
                impersonateHeader['CallerObjectId'] = impersonateAnotherUserInput;
            } else {
                alert('This should not happen. Wrong value in impersonateAnotherUserSelect: ' + impersonateAnotherUserField);
                return;
            }
        }

        const submitLink = document.getElementsByClassName('mainPanel')[0].getElementsByClassName('submitLink')[0];

        if (!!previewChangesBeforeSaving) {
            previewChanges(changedFields, pluralName, id, impersonateHeader);
            submitLink.style.display = 'none';
        } else {
            await commitSave(pluralName, id, changedFields, impersonateHeader);
        }
    }

    async function commitSave(pluralName, id, changedFields, impersonateHeader) {
        const requestUrl = apiUrl + pluralName + '(' + id + ')';

        Object.keys(changedFields)
            .filter(key => key.includes('____lookupOverride'))
            .forEach(key => delete changedFields[key]);

        let headers = {
            'accept': 'application/json',
            'content-type': 'application/json',
            'OData-MaxVersion': '4.0',
            'OData-Version': '4.0',
            'If-Match': '*'
        };

        headers = { ...headers, ...impersonateHeader }

        const bypassCustomPluginExecution = document.getElementsByClassName('bypassPluginExecutionBox')[0].checked;
        if (!!bypassCustomPluginExecution) {
            headers['MSCRM.BypassCustomPluginExecution'] = true;
        }

        const response = await fetch(requestUrl, {
            method: 'PATCH',
            headers: headers,
            body: JSON.stringify(changedFields)
        });

        if (response.ok) {
            await makeItPretty();
        } else {
            const errorText = await response.text();
            console.error(`${response.status} - ${errorText}`);
            window.alert(`${response.status} - ${errorText}`);
        }
    }

    function stringContains(str, value) {
        return str.indexOf(value) !== -1;
    }

    function isAnnotation(key) {
        return stringContains(key, formattedValueType) || stringContains(key, navigationPropertyType) || stringContains(key, lookupType);
    }

    async function prettifyWebApi(jsonObj, htmlElement, pluralName, generateEditLink) {
        const isMultiple = (jsonObj.value && Array.isArray(jsonObj.value));

        const result = await retrieveLogicalNameFromPluralNameAsync(pluralName);

        if (isMultiple) {
            const valueKeyWithCount = 'value (' + jsonObj.value.length + ' records)';

            jsonObj[valueKeyWithCount] = jsonObj.value;
            delete jsonObj.value;

            for (const key in jsonObj[valueKeyWithCount]) {
                jsonObj[valueKeyWithCount][key] = await enrichObjectWithHtml(jsonObj[valueKeyWithCount][key], result.logicalName, pluralName, result.primaryIdAttribute, false, false, 2);
            }
        } else {
            if (generateEditLink) {
                window.originalResponseCopy = JSON.parse(JSON.stringify(jsonObj));
            }
            jsonObj = await enrichObjectWithHtml(jsonObj, result.logicalName, pluralName, result.primaryIdAttribute, generateEditLink, false, 1);
        }

        let json = JSON.stringify(jsonObj, undefined, 3);
        json = json.replaceAll('"', '').replaceAll(replacedQuote, escapeHtml('"'));
        json = json.replaceAll(',', '').replaceAll(replacedComma, ',');

        htmlElement.innerText = '';
        const pre = document.createElement('pre');

        if (generateEditLink) {
            pre.classList.add('mainPanel');
        }

        htmlElement.appendChild(pre).innerHTML = json;
        setPreviewLinkClickHandlers();
        setEditLinkClickHandlers();
        setCopyToClipboardHandlers();
        setLookupEditHandlers();

        if (!isMultiple && generateEditLink) {
            setImpersonateUserHandlers();
        }
    }

    function previewChanges(changedFields, pluralName, id, impersonateHeader) {
        const changes = [];

        for (let key in changedFields) {
            const change = {};
            let originalValue = window.originalResponseCopy[key];
            const updatedValue = changedFields[key];

            if (changedFields[key + '____lookupOverride'] === true) {
                originalValue = changedFields[key + '____lookupOverrideOriginalValue'];

                delete changedFields[key + '____lookupOverride'];
                delete changedFields[key + '____lookupOverrideOriginalValue'];
            }

            if (key.includes('____lookupOverride')) {
                continue;
            }

            change.column = key.replace('@odata.bind', '');
            change.old = originalValue;
            change.new = updatedValue;
            changes.push(change);
        }

        const table = tableFromChanges(changes);

        // disable all stuff to prevent edits after previewing
        disableAllInputs();

        const editMenu = document.getElementById('previewChangesDiv');
        editMenu.innerHTML = '  ';
        editMenu.appendChild(table);

        const lineBreak = document.createElement('br');
        editMenu.appendChild(lineBreak);
        editMenu.append('    ');

        const undoAllLink = document.createElement('a');
        undoAllLink.innerText = 'Cancel';
        undoAllLink.href = 'javascript:';

        undoAllLink.onclick = makeItPretty;

        editMenu.appendChild(undoAllLink);

        const lineBreak2 = document.createElement('br');
        editMenu.appendChild(lineBreak2);
        editMenu.append('    ');

        const submitChangesLink = document.createElement('a');
        submitChangesLink.innerText = 'Commit Save';
        submitChangesLink.href = 'javascript:';

        // create this callback so we enclose the values we need when saving
        const saveCallback = async function () {
            await commitSave(pluralName, id, changedFields, impersonateHeader);
        }

        submitChangesLink.onclick = saveCallback;

        editMenu.appendChild(submitChangesLink);
    }

    function disableAllInputs() {
        const inputs = document.getElementsByTagName('input');
        for (let i = 0; i < inputs.length; i++) {
            inputs[i].disabled = true;
        }
        const selects = document.getElementsByTagName('select');
        for (let i = 0; i < selects.length; i++) {
            selects[i].disabled = true;
        }
        const textareas = document.getElementsByTagName('textarea');
        for (let i = 0; i < textareas.length; i++) {
            textareas[i].disabled = true;
        }
        const lookupEditLinks = document.getElementsByClassName('lookupEditLinks');
        for (let i = 0; i < lookupEditLinks.length; i++) {
            lookupEditLinks[i].style.display = 'none';
        }
    }

    async function previewRecord(pluralName, url) {
        const cssBody = `body {
            display: inline-flex;
            margin-top: 0px;
            margin-bottom: 0px;
            }
            `
        addcss(cssBody);
        const cssPre = `pre {
            width: 49vw;
            overflow-x: scroll;
            overflow-y: scroll;
            height: 100%;
            margin: 0px;
            }
            `
        addcss(cssPre);

        const newDiv = document.createElement('div');
        newDiv.classList.add('previewPanel');
        document.body.appendChild(newDiv);

        newDiv.style = 'position:relative;'

        const response = await odataFetch(url);

        await prettifyWebApi(response, newDiv, pluralName, false);

        const btn = document.createElement('button');
        btn.style = `
            height: 30px;
            width: auto;
            margin-right: 24px;
            margin-top: 10px;
            position: absolute;
            right: 10px;
            cursor: pointer;
            padding:0;
            font-size:24;
            padding: 0px 4px 0px 4px;
            `

        btn.innerHTML = '<div>Close Preview</div>';

        btn.addEventListener('click', function () {
            if (document.getElementsByClassName('previewPanel').length === 1) {
                resetCSS();
            }

            newDiv.remove();
        });

        newDiv.firstChild.insertBefore(btn, newDiv.firstChild.firstChild);

        newDiv.scrollIntoView();
    }

    function addHeaders(table, keys) {
        const header = table.createTHead();
        const row = header.insertRow(0);
        for (let i = 0; i < keys.length; i++) {
            const cell = row.insertCell();
            cell.appendChild(document.createTextNode(keys[i]));
        }
    }

    function tableFromChanges(changes) {
        const table = document.createElement('table');
        table.id = 'previewTable';

        if (changes.length === 0) {
            return table;
        }

        // create the table body
        for (let i = 0; i < changes.length; i++) {
            const change = changes[i];
            const row = table.insertRow();
            Object.keys(change).forEach(function (k) {
                const cell = row.insertCell();
                let value = change[k];

                if (value === null) {
                    value = '(empty)';
                }
                cell.appendChild(document.createTextNode(value));
            })
        }

        // add the header last to prevent issues with the body going into the header when the body is empty
        const header = changes[0];
        addHeaders(table, Object.keys(header));

        return table;
    }

    async function makeItPretty() {
        const response = await odataFetch(window.location.href);

        window.currentEntityPluralName = window.location.pathname.split('/').pop().split('(').shift();

        resetCSS();

        await prettifyWebApi(response, document.body, window.currentEntityPluralName, true);
    }

    function resetCSS() {
        const head = document.getElementsByTagName('head')[0];
        head.innerHTML = '';

        const css = `pre
            .string { color: brown; }
            .number { color: darkgreen; }
            .boolean { color: blue; }
            .null { color: magenta; }
            .guid { color: brown; }
            .link { color: blue; }
            .primarykey { color: tomato; }

            input {
                width: 300px;
                margin: 0 0 0 8px;
            }

            textarea {
                width: 400px;
                margin: 0 0 0 20px;
            }

            select {
                margin: 0 0 0 8px;
            }

            span:not(.lookupField):not(.lookupEdit):not(.link) {
                margin-right: 24px;
                padding-right: 16px;
            }

            option:empty {
                display:none;
            }

            .copyButton {
                color:dimgray;
                display: none;
            }           
            
            .field:hover .copyButton {
                display: unset;
            }

            .copyButton:field {
                color: black;
                cursor: pointer;
            }

            .copyButton:active {
                color: green;
            }             

            .link {
                margin:0;
                padding:0;
            }

            table {
                color: black;
                margin-left: 26px;
                border-collapse: collapse;
                border: 1px solid black;
                table-layout: fixed;
                width: 98%;
            }

            thead td {
                font-weight: bold;
            }

            td {
                padding: 4px;
                border: 1px solid;
                overflow:auto;
            }

            .impersonationIdFieldLabel {
                padding:0px;
                margin:0px;
            }
            `

        addcss(css);
    }

    await makeItPretty();
})()