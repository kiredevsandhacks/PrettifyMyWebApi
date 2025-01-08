(async function () {
    if (window.location.hash !== '#p' && window.location.hash !== '#pf' && window.location.hash !== '#pr') {
        return;
    }

    const formattedValueType = '@OData.Community.Display.V1.FormattedValue';
    const navigationPropertyType = '@Microsoft.Dynamics.CRM.associatednavigationproperty';
    const lookupType = '@Microsoft.Dynamics.CRM.lookuplogicalname';
    const metaDataKey = "@Microsoft.Dynamics.CRM.globalmetadataversion";
    const contextKey = "@odata.context";
    const etagKey = "@odata.etag";
    const totalRecordCountKey = "@Microsoft.Dynamics.CRM.totalrecordcount";
    const totalRecordCountLimitExceededKey = "@Microsoft.Dynamics.CRM.totalrecordcountlimitexceeded";

    const replacedQuote = '__~~__REPLACEDQUOTE__~~__';
    const replacedComma = '__~~__REPLACEDCOMMA__~~__';

    const clipBoardIcon = `<svg class='copyIcon' style='width:16px;position:absolute;padding-left:4px' viewBox='0 0 24 24'>
        <path fill='currentColor' d='M19 2h-4.18C14.4.84 13.3 0 12 0S9.6.84 9.18 2H5c-1.1 0-2 .9-2 2v16c0 1.1.9 2 2 2h14c1.1 0 2-.9 2-2V4c0-1.1-.9-2-2-2zm-7 0c.55 0 1 .45 1 1s-.45 1-1 1s-1-.45-1-1s.45-1 1-1zm7 18H5V4h2v3h10V4h2v16z'/>
    </svg>`.replaceAll(',', replacedComma); // need to 'escape' the commas because they cause issues with the JSON string cleanup code 

    const clipBoardIconCopied = `<svg class='copiedIcon' style='width:16px;position:absolute;display:none;padding-left:4px' viewBox='0 0 24 24'>
        <path fill='currentColor' d='M19 3h-4.18C14.4 1.84 13.3 1 12 1c-1.3 0-2.4.84-2.82 2H5c-1.1 0-2 .9-2 2v14c0 1.1.9 2 2 2h14c1.1 0 2-.9 2-2V5c0-1.1-.9-2-2-2zm-7 0c.55 0 1 .45 1 1s-.45 1-1 1s-1-.45-1-1s.45-1 1-1zm-2 14l-4-4l1.41-1.41L10 14.17l6.59-6.59L18 9l-8 8z'/>
    </svg>`.replaceAll(',', replacedComma); // need to 'escape' the commas because they cause issues with the JSON string cleanup code 

    const copyControlsHtml = `<span class='copyButton'>` + clipBoardIcon + `</span>` + `<span class='copiedNotification' style='display:none'>` + clipBoardIconCopied + `<span style='padding-left:24px'>Copied!</span></span>`;

    let apiUrl = '';
    let titleSet = false;

    try {
        apiUrl = /([\/a-zA-Z0-9]+)?\/api\/data\/v[0-9][0-9]?.[0-9]\//.exec(window.location.pathname)[0];
    } catch {
        alert('It seems you are not viewing a form or the dataverse odata web api. If you think this is an error, please contact the author of the extension and he will fix it asap.');
        return;
    }

    const lazyLookupInitFunctions = {};

    const memoizedCalls = {};

    const foundNavigationProperties = [];

    let dataTypeFilters = [];

    async function odataFetch(url, memoize) {
        if (memoize && memoizedCalls.hasOwnProperty(url)) {
            return memoizedCalls[url];
        }

        const response = await fetch(url, { headers: { 'Prefer': 'odata.include-annotations="*"', 'Cache-Control': 'no-cache' } });

        const responseJson = await response.json();

        if (memoize) {
            memoizedCalls[url] = responseJson;
        }

        return responseJson;
    }


    async function odataPatchWithErrorHandling(url, body) {
        let headers = {
            'accept': 'application/json',
            'content-type': 'application/json',
            'OData-MaxVersion': '4.0',
            'OData-Version': '4.0',
            'If-Match': '*',
        };

        const response = await fetch(url, { method: 'PATCH', headers: headers, body: JSON.stringify(body) });

        if (!response.ok) {
            const errorText = await response.text();
            console.error(`${response.status} - ${errorText}`);
            window.alert(`${response.status} - ${errorText}`);

            return false;
        }
        return true;
    }

    async function odataDeleteWithErrorHandling(url) {
        const response = await fetch(url, { method: 'DELETE', });

        if (!response.ok) {
            const errorText = await response.text();
            console.error(`${response.status} - ${errorText}`);
            window.alert(`${response.status} - ${errorText}`);

            return false;
        }
        return true;
    }

    async function retrievePluralName(logicalName) {
        const pluralNames = await retrieveAllPluralNames();

        const table = pluralNames.find(p => p.LogicalName === logicalName);

        return table.EntitySetName;
    }

    async function retrieveLogicalName(pluralName) {
        const pluralNames = await retrieveAllPluralNames();

        const table = pluralNames.find(p => p.EntitySetName === pluralName);

        return table.LogicalName;
    }

    async function retrieveAllPluralNames() {
        const requestUrl = apiUrl + "EntityDefinitions?$select=EntitySetName,LogicalName";

        const json = await odataFetch(requestUrl, true);

        return json.value;
    }

    async function retrieveRelationshipDefinition(relationShipName) {
        const relationships = await retrieveRelationships();
        const combined = [...relationships.ManyToManyRelationships, ...relationships.ManyToOneRelationships, ...relationships.OneToManyRelationships];
        const relationship = combined.filter(r =>
            r.SchemaName === relationShipName ||
            r.Entity1NavigationPropertyName === relationShipName ||
            r.Entity2NavigationPropertyName === relationShipName ||
            r.ReferencedEntityNavigationPropertyName === relationShipName ||
            r.ReferencingEntityNavigationPropertyName === relationShipName
        );

        if (relationship.length !== 1) {
            return null;
        }

        return relationship[0];
    }

    // TODO: if lookup query is implemented, maybe refactor this into the retrieveAllPluralNames function? Let it just retrieve all needed metadata from the system in one api call
    async function retrievePrimaryNameAndKeyAndPluralName(logicalName) {
        const requestUrl = apiUrl + "EntityDefinitions?$select=PrimaryNameAttribute,PrimaryIdAttribute,EntitySetName&$filter=(LogicalName eq '" + logicalName + "')";

        const json = await odataFetch(requestUrl, true);

        let primaryNameAndKey = {};
        primaryNameAndKey.name = json.value[0].PrimaryNameAttribute;
        primaryNameAndKey.key = json.value[0].PrimaryIdAttribute;
        primaryNameAndKey.plural = json.value[0].EntitySetName;

        return primaryNameAndKey;
    }

    async function retrieveLogicalAndPrimaryKeyAndPrimaryName(pluralName) {
        const requestUrl = apiUrl + "EntityDefinitions?$select=LogicalName,PrimaryIdAttribute,PrimaryNameAttribute&$filter=(EntitySetName eq '" + pluralName + "')";

        const json = await odataFetch(requestUrl, true);

        if (json.value.length === 0) {
            return {};
        }

        const logicalName = json.value[0].LogicalName;
        const primaryIdAttribute = json.value[0].PrimaryIdAttribute;
        const primaryNameAttribute = json.value[0].PrimaryNameAttribute;

        const result = {
            logicalName: logicalName,
            primaryIdAttribute: primaryIdAttribute,
            primaryNameAttribute: primaryNameAttribute
        };

        return result;
    }

    async function retrieveAttributesWithTypes(logicalName) {
        const requestUrl = apiUrl + `EntityDefinitions(LogicalName='${logicalName}')/Attributes?$select=LogicalName,AttributeType`;

        const json = await odataFetch(requestUrl, true);

        return json.value;
    }


    async function retrieveSearchableAttributes(logicalName) {
        const requestUrl = apiUrl + "EntityDefinitions(LogicalName='" + logicalName + "')/Attributes?$filter=IsValidForAdvancedFind/Value eq true";

        const json = await odataFetch(requestUrl, true);

        return json.value;
    }

    async function retrieveUpdateableAttributes(logicalName) {
        const requestUrl = apiUrl + "EntityDefinitions(LogicalName='" + logicalName + "')/Attributes?$filter=IsValidForUpdate eq true";

        const json = await odataFetch(requestUrl);

        return json.value;
    }

    async function retrieveCreatableAttributes(logicalName) {
        const requestUrl = apiUrl + "EntityDefinitions(LogicalName='" + logicalName + "')/Attributes?$filter=IsValidForCreate eq true";

        const json = await odataFetch(requestUrl, true);

        return json.value;
    }

    async function retrieveOptionSetMetadata(logicalName) {
        const requestUrl = apiUrl + "EntityDefinitions(LogicalName='" + logicalName + "')/Attributes/Microsoft.Dynamics.CRM.PicklistAttributeMetadata?$select=LogicalName&$expand=OptionSet,GlobalOptionSet";

        const json = await odataFetch(requestUrl, true);

        return json.value;
    }

    async function retrieveMultiSelectOptionSetMetadata(logicalName) {
        const requestUrl = apiUrl + "EntityDefinitions(LogicalName='" + logicalName + "')/Attributes/Microsoft.Dynamics.CRM.MultiSelectPicklistAttributeMetadata?$select=LogicalName&$expand=OptionSet,GlobalOptionSet";

        const json = await odataFetch(requestUrl, true);

        return json.value;
    }


    async function retrieveStateMetadata(logicalName) {
        const requestUrl = apiUrl + "EntityDefinitions(LogicalName='" + logicalName + "')/Attributes/Microsoft.Dynamics.CRM.StateAttributeMetadata?$select=LogicalName&$expand=OptionSet,GlobalOptionSet";

        const json = await odataFetch(requestUrl, true);

        return json.value;
    }

    async function retrieveStatusMetadata(logicalName) {
        const requestUrl = apiUrl + "EntityDefinitions(LogicalName='" + logicalName + "')/Attributes/Microsoft.Dynamics.CRM.StatusAttributeMetadata?$select=LogicalName&$expand=OptionSet,GlobalOptionSet";

        const json = await odataFetch(requestUrl, true);

        return json.value;
    }

    async function retrieveBooleanFieldMetadata(logicalName) {
        const requestUrl = apiUrl + "EntityDefinitions(LogicalName='" + logicalName + "')/Attributes/Microsoft.Dynamics.CRM.BooleanAttributeMetadata?$select=LogicalName&$expand=OptionSet,GlobalOptionSet";

        const json = await odataFetch(requestUrl, true);

        return json.value;
    }

    async function initEntityClientMetadataForCurrentRecord() {
        const currentRecordLogicalName = await retrieveLogicalName(window.currentEntityPluralName);
        await retrieveEntityClientMetadata(currentRecordLogicalName);
    }

    async function retrieveEntityClientMetadata(logicalName) {
        const requestUrl = apiUrl + "GetClientMetadata(ClientMetadataQuery=@ClientMetadataQuery)?@ClientMetadataQuery={'MetadataType':'entity','MetadataSubtype':'{\"" + logicalName + "\":[\"merged\"]}'}";

        const json = await odataFetch(requestUrl, true);
        const clientMetadata = JSON.parse(json.Metadata);

        return clientMetadata;
    }

    async function retrieveNavigationProperty(lookupEntitylogicalName, fieldName) {
        const currentRecordLogicalName = await retrieveLogicalName(window.currentEntityPluralName);

        const metadata = await retrieveEntityClientMetadata(currentRecordLogicalName);

        if (metadata.Entities.length != 1) {
            alert(`Something went wrong with retrieving the navigation property for ${currentRecordLogicalName}/${lookupEntitylogicalName}/${fieldName}: ${metadata.Entities.length} entities found. This should not happen.`);
            return null;
        }

        let relations = [];

        if (fieldName === 'ownerid') {
            relations = metadata.Entities[0].ManyToOneRelationships.filter(rel => rel.ReferencingAttribute === fieldName);
        } else {
            relations = metadata.Entities[0].ManyToOneRelationships.filter(rel => rel.ReferencingEntity === currentRecordLogicalName && rel.ReferencedEntity === lookupEntitylogicalName && rel.ReferencingAttribute === fieldName);
        }

        if (relations.length != 1) {
            alert(`Something went wrong with retrieving the navigation property for ${currentRecordLogicalName}/${lookupEntitylogicalName}/${fieldName}: ${relations.length} relations found. This should not happen.`);
            return null;
        }

        return relations[0].ReferencingEntityNavigationPropertyName;
    }

    async function retrieveRelationships() {
        const currentRecordLogicalName = await retrieveLogicalName(window.currentEntityPluralName);

        const metadata = await retrieveEntityClientMetadata(currentRecordLogicalName);

        if (metadata.Entities.length != 1) {
            return null;
        }

        const result = {
            ManyToManyRelationships: metadata.Entities[0].ManyToManyRelationships,
            OneToManyRelationships: metadata.Entities[0].OneToManyRelationships,
            ManyToOneRelationships: metadata.Entities[0].ManyToOneRelationships
        };

        return result;
    }

    async function doLookupInputQuery(pluralName, fieldToFilterOn, fieldToDisplay, primaryKeyfieldname, query) {
        query = query.replaceAll(`'`, `''`);
        const requestUrl = apiUrl + `${pluralName}?$top=10&$select=${fieldToDisplay},${primaryKeyfieldname}&$filter=contains(${fieldToFilterOn}, '${query}')`;

        const json = await odataFetch(requestUrl, false);

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
        const newLocation = apiUrl + escapeHtml(pluralName) + '(' + escapeHtml(formattedGuid) + ')#p';

        return `<a target='_blank' href='${newLocation}'>${escapeHtml(formattedGuid)}</a>`;
    }

    function generateFormUrlAnchor(logicalName, guid) {
        const newLocation = '/main.aspx?etn=' + escapeHtml(logicalName) + '&id=' + escapeHtml(guid) + '&pagetype=entityrecord';

        return `<a target='_blank' href='${newLocation}'>Open in Form</a>`;
    }

    function generateWebApiAnchor(guid, pluralName) {
        const formattedGuid = guid.replace('{', '').replace('}', '');
        const newLocation = apiUrl + escapeHtml(pluralName) + '(' + escapeHtml(formattedGuid) + ')#p';

        return `<a target='_blank' href='${newLocation}'>Open in Web Api</a>`;
    }

    function generateDisassociateAnchor(pluralName, guid, navigationProperty, classAddition) {
        const formattedGuid = guid.replace('{', '').replace('}', '');

        return `<a class='disassociateLink${escapeHtml(classAddition)}' data-pluralName='${escapeHtml(pluralName)}' data-guid='${escapeHtml(formattedGuid)}' data-navigationproperty='${escapeHtml(navigationProperty)}' href='javascript:'>Disassociate this row</a>`;
    }

    async function generatePreviewUrlAnchor(logicalName, guid) {
        const pluralName = await retrievePluralName(logicalName);
        const formattedGuid = guid.replace('{', '').replace('}', '');

        return `<a class='previewLink' data-pluralName='${escapeHtml(pluralName)}' data-guid='${escapeHtml(formattedGuid)}' href='javascript:'>Preview</a>`;
    }

    async function generateDeleteAnchor(logicalName, guid) {
        const pluralName = await retrievePluralName(logicalName);
        const formattedGuid = guid.replace('{', '').replace('}', '');

        return `<a class='deleteLink' data-logicalName='${escapeHtml(logicalName)}' data-pluralName='${escapeHtml(pluralName)}' data-guid='${escapeHtml(formattedGuid)}' href='javascript:'>Delete this row</a>`
    }

    async function generateEditMenu(logicalName, guid, isCreateMode) {
        const pluralName = await retrievePluralName(logicalName);
        let formattedGuid = guid?.replace('{', '')?.replace('}', '');

        if (isCreateMode) {
            formattedGuid = '{temp}'
        }

        return `
<a class='editLink' data-logicalName='${escapeHtml(logicalName)}' data-pluralName='${escapeHtml(pluralName)}' data-guid='${escapeHtml(formattedGuid)}' href='javascript:'>Edit this row</a>     
    <div class='editMenuDiv' style='display: none;'>
        <div class='checkBoxDiv'>    Bypass Custom Plugin execution<input class='bypassPluginExecutionBox' type='checkbox' style='width:25px;'>
        </div><div class='checkBoxDiv'>    Bypass Power Automate Flow execution<input class='bypassFlowExecutionBox' type='checkbox' style='width:25px;'>
        </div><div class='checkBoxDiv'>    Preview changes before committing save<input class='previewChangesBeforeSavingBox' type='checkbox' style='width:25px;' checked='true'>
        </div><div class='checkBoxDiv'>    Impersonate another user<input class='impersonateAnotherUserCheckbox' type='checkbox' style='width:25px;'>
        </div><div class='checkBoxDiv'>    Remember these settings (for this environment)<input id='saveSettingsInput' type='checkbox' style='width:25px;'>
        </div><div class='impersonateDiv' style='display:none;'><div>      Base impersonation on this field: <select  class='impersonateAnotherUserSelect'><option value='systemuserid'>systemuserid</option><option value='azureactivedirectoryobjectid'>azureactivedirectoryobjectid</option></select>  <i><a href='https://learn.microsoft.com/en-us/power-apps/developer/data-platform/webapi/impersonate-another-user-web-api#how-to-impersonate-a-user' target='_blank'>What's this?</a></i>
        </div><div>      <span class='impersonationIdFieldLabel'>systemuserid:</span><input class='impersonateAnotherUserInput' placeholder='00000000-0000-0000-0000-000000000000'>  <span class='impersonateUserPreview'></span>
        </div></div><div><div id='previewChangesDiv'></div>    <a class='cancelLink' href='javascript:'>Cancel</a><br/>    <a class='submitLink' style='display: none;' href='javascript:'>Save</a>
        <div class='saveInProgressDiv' style='display:none;' >    Saving...</div>
        </div>
    </div>`.replaceAll('\n', '');
    }

    function createSpan(cls, value) {
        return `<span style='display: inline-flex;' class='${escapeHtml(cls)} field'>${escapeHtml(value)}${copyControlsHtml}</span>`;
    }

    function createSpanForLookup(cls, value) {
        return `<span style='display: inline-flex;' class='${escapeHtml(cls)} field lookupField'>${escapeHtml(value)}${copyControlsHtml}</span>`;
    }

    function createLinkSpan(cls, value) {
        // unsafe contents of 'value' have been escaped in a previous stage
        return `<span class='${escapeHtml(cls)}'>${value}</span>`;
    }

    function createFieldSpan(cls, value, fieldName) {
        return `<span style='display: inline-flex;' class='${escapeHtml(cls)} field'>${escapeHtml(value)}<div class='inputContainer containerNotEnabled' style='display: none;' data-fieldname='${escapeHtml(fieldName)}'></div>${copyControlsHtml}</span>`;
    }

    function createLookupEditField(displayName, guid, fieldname, lookupTypeValue) {
        if (displayName == null) {
            displayName = '';
        }

        fieldname = fieldname.substring(1, fieldname.length - 1).substring(0, fieldname.length - 7);

        const formattedGuid = guid?.replace('{', '')?.replace('}', '');
        return `<div class='lookupEditLinks' style='display:none;' data-fieldname='${escapeHtml(fieldname)}'><span class='pf-link'>   <a href='javascript:' class='searchDifferentRecord lookupEditLink' data-fieldname='${escapeHtml(fieldname)}'>Edit lookup</a></span><span class='pf-link'>   <a href='javascript:' class='clearLookup lookupEditLink' data-fieldname='${escapeHtml(fieldname)}'>Clear lookup</a></span><span class='pf-link'>   <a href='javascript:' class='cancelLookupEdit lookupEditLink' data-fieldname='${escapeHtml(fieldname)}' style='display:none;'>Undo changes</a></span></div><span class='lookupEdit' style='display: none;'><div class='inputContainer containerNotEnabled' data-name='${escapeHtml(displayName)}' data-id='${escapeHtml(formattedGuid)}' data-fieldname='${escapeHtml(fieldname)}' data-lookuptype='${escapeHtml(lookupTypeValue)}'></div></span>`;
    }

    function createdFormattedValueSpan(cls, value, fieldName, formattedValue) {
        let insertedValue = '';

        // toString the value because it can be a number. The formattedValue is always a string
        if (value?.toString() !== formattedValue) {
            insertedValue = value + ' : ' + formattedValue;
        } else {
            insertedValue = value;
        }

        return `<span style='display: inline-flex;' class='${escapeHtml(cls)} field'>${escapeHtml(insertedValue)}<div class='inputContainer containerNotEnabled' style='display: none;' data-fieldname='${escapeHtml(fieldName)}'></div>${copyControlsHtml}</span>`;
    }

    async function enrichObjectWithHtml(jsonObj, logicalName, pluralName, primaryIdAttribute, isSingleRecord, isNested, nestedLevel, primaryNameAttribute, isCreateMode, relationShipDefinition, isSingleColumnValueOnly) {
        const recordId = jsonObj[primaryIdAttribute]; // we need to get this value before parsing or else it will contain html

        const ordered = orderProperties(jsonObj);

        for (let key in ordered) {
            let value = ordered[key];

            const cls = determineType(value);

            if (key === contextKey || key === metaDataKey || key === etagKey) {
                // get outta here
                delete ordered[key];
                continue;
            }

            let attributeType = '';
            if (logicalName != null) {
                const keyForDataTypeFilter = key.split('@')[0];
                const attributesWithTypes = await retrieveAttributesWithTypes(logicalName);
                const attribute = attributesWithTypes.filter(a => a.LogicalName === keyForDataTypeFilter || '_' + a.LogicalName + '_value' === keyForDataTypeFilter)[0];
                if (attribute != null) {
                    attributeType = attribute.AttributeType;
                }
            }

            if (Array.isArray(value)) {
                ordered[key] = [];

                if (Object.values(value).every(v => typeof (v) === 'object')) {
                    for (let nestedKey in value) {
                        let nestedValue = value[nestedKey];

                        ordered[key][nestedKey] = await enrichObjectWithHtml(nestedValue, null, null, null, null, true, nestedLevel + 1, null, false, null, false);
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
                ordered[key] = await enrichObjectWithHtml(value, null, null, null, null, true, nestedLevel + 1, null, false, null, false);
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

                foundNavigationProperties.push(navigationPropertyValue);

                const newApiUrl = await generateApiAnchorAsync(lookupTypeValue, value);
                const formUrl = generateFormUrlAnchor(lookupTypeValue, value);
                const previewUrl = await generatePreviewUrlAnchor(lookupTypeValue, value);

                let lookupFormatted = '';

                lookupFormatted += `<span class='lookupDisplay'>{<br>   ` + spaces +
                    createLinkSpan('pf-link', newApiUrl) + ' - ' +
                    createLinkSpan('pf-link', formUrl) + ' - ' +
                    createLinkSpan('pf-link', previewUrl);
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
                lookupFormatted += createLookupEditField(formattedValueValue, value, key, lookupTypeValue);
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
                let formattedValueValue = ordered[key + formattedValueType];

                formattedValueValue = formattedValueValue.replaceAll(replacedQuote, ''); // to prevent malformed html and potential xss, disallow this string
                formattedValueValue = formattedValueValue.replaceAll('"', replacedQuote);
                formattedValueValue = formattedValueValue.replaceAll(',', replacedComma);

                ordered[key] = createdFormattedValueSpan(cls, value, key, formattedValueValue);
                delete ordered[key + formattedValueType];
            } else {
                if (logicalName !== 'solution' && key === 'solutionid' && value != null) {
                    ordered[key] = createFieldSpan(cls, value, key) + generateWebApiAnchor(value, 'solutions');
                }
                else if (key === primaryIdAttribute) {
                    ordered[key] = '<b>' + createSpan('primarykey', value) + '</b>';
                } else if (key === primaryNameAttribute) {
                    ordered[key] = '<b>' + createFieldSpan(cls, value, key) + '</b>';
                } else {
                    ordered[key] = createFieldSpan(cls, value, key);
                }
            }

            // put the key inside an html element for easier handling later on
            const keyCopy = ordered[key];
            delete ordered[key];
            const escapedKey = escapeHtml(key);
            const keySpan = `<span class='keySpan' data-cls='${cls}' data-attributetype='${escapeHtml(attributeType)}' data-key='${escapedKey}'>${escapedKey}</span>`
            ordered[keySpan] = keyCopy;
        }

        const newObj = {};
        if (!isNested) {
            if (logicalName != null && logicalName !== '' && recordId != null && recordId !== '') {
                newObj['Form Link'] = createLinkSpan('pf-link', generateFormUrlAnchor(logicalName, recordId));

                if (isSingleRecord) {
                    newObj['Edit this row'] = createLinkSpan('pf-link', await generateEditMenu(logicalName, recordId, false));
                    newObj['Delete this row'] = createLinkSpan('pf-link', await generateDeleteAnchor(logicalName, recordId, isSingleRecord));
                } else {
                    newObj['Web Api Link'] = createLinkSpan('pf-link', generateWebApiAnchor(recordId, pluralName));
                }

                if (relationShipDefinition != null) {
                    if (window.isManyToOneContext) {
                        newObj['Disassociate this row'] = createLinkSpan('pf-link', generateDisassociateAnchor(pluralName, recordId, relationShipDefinition.ReferencingEntityNavigationPropertyName, 'ManyToOne'));
                    } else if (relationShipDefinition.RelationshipType === 1) {
                        newObj['Disassociate this row'] = createLinkSpan('pf-link', generateDisassociateAnchor(pluralName, recordId, '', 'ManyToMany'));
                    } else if (relationShipDefinition.RelationshipType === 0) {
                        newObj['Disassociate this row'] = createLinkSpan('pf-link', generateDisassociateAnchor(pluralName, recordId, relationShipDefinition.ReferencingEntityNavigationPropertyName, 'OneToMany',));
                    }
                }
            } else if (!isSingleColumnValueOnly && !isCreateMode && logicalName != null && logicalName !== '' && (recordId == null || recordId === '')) {
                newObj['Form Link'] = 'Could not generate link';
                newObj['Web Api Link'] = 'Could not generate link';
            }
        }

        if (isCreateMode) {
            newObj['Create new row'] = createLinkSpan('pf-link', await generateEditMenu(logicalName, recordId, true));
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
        Array.from(document.querySelectorAll('.field')).forEach((el) => el.onclick = (element) => {
            if (window.pfwaMode !== 'read') {
                return;
            }

            // to prevent weird issues with innerText and double clicking
            if (window.copyToClipBoardBusy) {
                return;
            }

            try {
                window.copyToClipBoardBusy = true;
                navigator.clipboard.writeText(el.innerText).then(() => {
                    const copyIcon = el.querySelector('.copyIcon');
                    const copiedIcon = el.querySelector('.copiedIcon');
                    const copiedNotification = el.querySelector('.copiedNotification');
                    copiedNotification.style.display = 'unset';
                    copiedIcon.style.display = 'unset';
                    copyIcon.style.display = 'none';

                    setTimeout(() => {
                        copiedNotification.style.display = 'none';
                        copiedIcon.style.display = 'none';
                        copyIcon.style.display = 'unset';

                        window.copyToClipBoardBusy = false;
                    }, 400);

                    console.log('Content copied to clipboard');
                }, () => {
                    window.copyToClipBoardBusy = false;
                    alert('Failed to copy');
                });
            } catch {
                window.copyToClipBoardBusy = false;
            }
        });
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

        for (let editLink of editLinks) {
            const logicalName = editLink.dataset.logicalname;
            const pluralName = editLink.dataset.pluralname;
            const id = editLink.dataset.guid;

            editLink.onclick = async function () {
                await editRecord(logicalName, pluralName, id, false);
            }
        }
    }

    function setDeleteLinkClickHandlers() {
        const deleteLinks = document.getElementsByClassName('deleteLink');

        for (let deleteLink of deleteLinks) {
            const logicalName = deleteLink.dataset.logicalname;
            const pluralName = deleteLink.dataset.pluralname;
            const id = deleteLink.dataset.guid;

            deleteLink.onclick = async function () {
                await deleteRecord(logicalName, pluralName, id);
            }
        }
    }

    function setDisassociateClickHandlers() {
        const disassociateLinksOneToMany = document.getElementsByClassName('disassociateLinkOneToMany');

        for (let disassociateLink of disassociateLinksOneToMany) {
            const pluralName = disassociateLink.dataset.pluralname;
            const id = disassociateLink.dataset.guid;
            const navigationProperty = disassociateLink.dataset.navigationproperty

            disassociateLink.onclick = async function () {
                disassociateLink.style.display = 'none';
                await disassociateRowOneToMany(pluralName, id, navigationProperty);
                disassociateLink.style.display = 'unset';
            }
        }

        const disassociateLinksManyToOne = document.getElementsByClassName('disassociateLinkManyToOne');

        for (let disassociateLink of disassociateLinksManyToOne) {
            const navigationProperty = disassociateLink.dataset.navigationproperty

            disassociateLink.onclick = async function () {
                disassociateLink.style.display = 'none';
                await disassociateRowManyToOne(navigationProperty);
                disassociateLink.style.display = 'unset';
            }
        }

        const disassociateLinksManyToMany = document.getElementsByClassName('disassociateLinkManyToMany');

        for (let disassociateLink of disassociateLinksManyToMany) {
            const pluralName = disassociateLink.dataset.pluralname;
            const id = disassociateLink.dataset.guid;

            disassociateLink.onclick = async function () {
                disassociateLink.style.display = 'none';
                await disassociateRowManyToMany(pluralName, id);
                disassociateLink.style.display = 'unset';
            }
        }
    }

    async function disassociateRowManyToOne(navigationProperty) {
        if (confirm('Are you sure you want to disassociate this row?')) {
            const body = {};
            body[navigationProperty + '@odata.bind'] = null;

            const recordId = location.pathname.split(')')[0].split('(')[1];

            const result = await odataPatchWithErrorHandling(apiUrl + window.originalEntityPluralName + '(' + recordId + ')', body);

            if (result) {
                alert('Row was disassociated. This page will now reload.');
                window.location.reload();
            }
        }
    }

    async function disassociateRowOneToMany(pluralName, id, navigationProperty) {
        if (confirm('Are you sure you want to disassociate this row?')) {
            const body = {};
            body[navigationProperty + '@odata.bind'] = null;

            const result = await odataPatchWithErrorHandling(apiUrl + pluralName + '(' + id + ')', body);

            if (result) {
                alert('Row was disassociated. This page will now reload.');
                window.location.reload();
            }
        }
    }

    async function disassociateRowManyToMany(pluralName, id, navigationProperty) {
        if (confirm('Are you sure you want to disassociate this row?')) {
            const url = window.location.pathname + '/$ref?$id=' + location.origin + apiUrl + pluralName + '(' + id + ')';
            const result = await odataDeleteWithErrorHandling(url);

            if (result) {
                alert('Row was disassociated. This page will now reload.');
                window.location.reload();
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
            const retrievedSystemUser = await odataFetch(apiUrl + `systemusers?$filter=${impersonateAnotherUserSelect.value} eq '${impersonateAnotherUserInput.value}'&$select=fullname`, false);
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

            const retrievedRecord = await odataFetch(apiUrl + `${plural}(${input.value})?$select=${name}`, false);

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
            const fieldName = link.dataset.fieldname;
            link.onclick = async () => {
                const enabledInputFields = document.getElementsByClassName('enabledInputField');
                const lookup = [...enabledInputFields].find(f => f.dataset.fieldname === fieldName);

                const navigationProperty = await retrieveNavigationProperty(lookup.dataset.logicalname, fieldName);

                if (navigationProperty == null) {
                    // user was already alerted at this point
                    return;
                }

                lookup.dataset.navigationproperty = navigationProperty;

                lookup.value = null;
                lookup.dataset.id = null;
                lookup.dataset.editmode = 'makenull';

                const makeNullDiv = [...document.getElementsByClassName('makeLookupNullDiv')].find(f => f.dataset.fieldname === fieldName);
                makeNullDiv.style.display = 'block';

                const lookupEditMenuDivs = document.getElementsByClassName('lookupEditMenuDiv');
                const lookupEditMenuDiv = [...lookupEditMenuDivs].find(f => f.dataset.fieldname === fieldName);

                lookupEditMenuDiv.style.display = 'none';

                const cancelLookupEditDivs = document.getElementsByClassName('cancelLookupEdit');
                const cancelLookupEditDiv = [...cancelLookupEditDivs].find(f => f.dataset.fieldname === fieldName);
                cancelLookupEditDiv.style.display = 'unset';
            };
        }

        const searchDifferentRecordLinks = document.getElementsByClassName('searchDifferentRecord');

        for (let link of searchDifferentRecordLinks) {
            const fieldName = link.dataset.fieldname;
            link.onclick = async () => {
                const enabledInputFields = document.getElementsByClassName('enabledInputField');
                const lookup = [...enabledInputFields].find(f => f.dataset.fieldname === fieldName);
                const navigationProperty = await retrieveNavigationProperty(lookup.dataset.logicalname, fieldName);

                if (navigationProperty == null) {
                    // user was already alerted at this point
                    return;
                }

                lookup.dataset.navigationproperty = navigationProperty;

                lookup.dataset.editmode = 'update';

                const lookupEditMenuDivs = document.getElementsByClassName('lookupEditMenuDiv');
                const lookupEditMenuDiv = [...lookupEditMenuDivs].find(f => f.dataset.fieldname === fieldName);

                lookupEditMenuDiv.style.display = 'unset';

                const makeNullDiv = [...document.getElementsByClassName('makeLookupNullDiv')].find(f => f.dataset.fieldname === fieldName);
                makeNullDiv.style.display = 'none';

                const cancelLookupEditDivs = document.getElementsByClassName('cancelLookupEdit');
                const cancelLookupEditDiv = [...cancelLookupEditDivs].find(f => f.dataset.fieldname === fieldName);
                cancelLookupEditDiv.style.display = 'unset';

                // uncommented for now
                // await lazyLookupInitFunctions[fieldName]?.call();
            };
        }

        const cancelLookupEditLinks = document.getElementsByClassName('cancelLookupEdit');

        for (let link of cancelLookupEditLinks) {
            const fieldName = link.dataset.fieldname;
            link.onclick = () => {
                const enabledInputFields = document.getElementsByClassName('enabledInputField');
                const lookup = [...enabledInputFields].find(f => f.dataset.fieldname === fieldName);
                lookup.value = lookup.dataset.originalname;
                lookup.dataset.id = lookup.dataset.originalid;
                lookup.dataset.editmode = null;

                link.style.display = 'none';

                const makeNullDiv = [...document.getElementsByClassName('makeLookupNullDiv')].find(f => f.dataset.fieldname === fieldName);
                makeNullDiv.style.display = 'none';

                const lookupEditMenuDivs = document.getElementsByClassName('lookupEditMenuDiv');
                const lookupEditMenuDiv = [...lookupEditMenuDivs].find(f => f.dataset.fieldname === fieldName);

                lookupEditMenuDiv.style.display = 'none';
            };
        }
    }

    function createInput(container, multiLine, datatype) {
        let value = null;
        if (window.pfwaMode === 'update') {
            value = window.originalResponseCopy[container.dataset.fieldname];
        }

        let input;

        if (!multiLine) {
            input = document.createElement('input');
        } else {
            input = document.createElement('textarea');
        }

        input.value = value;

        setInputMetadata(input, container, datatype);

        // in some very rare cases, dataverse will return an empty string instead of null 
        if (value === '') {
            input.dataset.isEmptyString = 'true';
        }
    }

    function createMultiSelectOptionSetValueInput(container, optionSet) {
        const values = window.originalResponseCopy[container.dataset.fieldname]?.split(',');

        const multiSelectDivContainer = document.createElement('div');
        const fakeSelect = document.createElement('select');
        fakeSelect.innerHTML = `<select style='border:none;outline:none;'><option style='display:none;'></option></select>`;

        const multiSelectDiv = document.createElement('div');
        multiSelectDiv.classList.add('multiSelectDiv')
        multiSelectDiv.style = 'display:none;'

        let multiSelectDivHtml = '';

        optionSet.forEach(function (option) {
            const formattedOption = option.Value + ' : ' + option.Label?.UserLocalizedLabel?.Label;

            const isSelected = values?.find(v => v === option.Value?.toString()) != null;

            const checked = isSelected ? 'checked' : '';

            multiSelectDivHtml += `<div class='multiSelectSubDiv'><input class='multiSelectInput' type='checkbox' ${checked} data-label='${escapeHtml(option.Label?.UserLocalizedLabel?.Label)}' data-value='${escapeHtml(option.Value)}'>${escapeHtml(formattedOption)}</div>`;
        });

        multiSelectDiv.innerHTML = multiSelectDivHtml;

        multiSelectDivContainer.appendChild(fakeSelect);
        multiSelectDivContainer.appendChild(multiSelectDiv);

        fakeSelect.onclick = (e) => {
            e.preventDefault();
            fakeSelect.blur();
            window.focus();

            multiSelectDiv.style.display = 'unset';
            transParentOverlay.style.display = 'unset';
        };

        const updateLabel = function () {
            let selectLabel = '';
            multiSelectDiv.querySelectorAll('input').forEach(input => {
                if (input.checked) {
                    if (selectLabel !== '') {
                        selectLabel += ', ';
                    }

                    selectLabel += input.dataset.label;
                }
            })

            if (selectLabel === '') {
                selectLabel = '(no options selected)';
            }

            fakeSelect.options[0].innerText = selectLabel;
        }

        updateLabel();

        multiSelectDiv.querySelectorAll('input').forEach(input => {
            input.onchange = updateLabel;
        });

        setInputMetadata(multiSelectDivContainer, container, 'multiselectoption');
    }

    function createOptionSetValueInput(container, optionSet, nullable, editable, isStatus) {
        let value = null;
        if (window.pfwaMode === 'update') {
            value = window.originalResponseCopy[container.dataset.fieldname];
        }
        const select = document.createElement('select');

        let selectHtml = "";

        if (!editable) {
            select.setAttribute('disabled', '');
        }

        if (nullable) {
            selectHtml = "<option value='null'>null</option>"; // empty option for clearing it
        }

        let mappedValue = null;

        optionSet.forEach(function (option) {
            let formattedOption = option.Value + ' : ' + option.Label?.UserLocalizedLabel?.Label;
            let formattedEscapedOption = escapeHtml(formattedOption);
            if (value === option.Value) {
                mappedValue = formattedOption;
            }

            // TODO: refactor the value attribute to contain the pure values, true/false/null
            if (option.State || option.State === 0) {
                selectHtml += `<option data-state='${escapeHtml(option.State)}' value='${formattedEscapedOption}'>${formattedEscapedOption}</option>`;
            } else {
                selectHtml += `<option value='${formattedEscapedOption}'>${formattedEscapedOption}</option>`;
            }
        });

        if (mappedValue == null && value != null) {
            let formattedOption = value + ' : (unmapped)';
            let formattedEscapedOption = escapeHtml(formattedOption);
            selectHtml += `<option data-state='${escapeHtml(value)}' value='${formattedEscapedOption}'>${formattedEscapedOption}</option>`;
            mappedValue = formattedOption;
        }

        select.innerHTML = selectHtml;

        if (mappedValue != null) {
            select.value = mappedValue;
        }

        setInputMetadata(select, container, 'option');

        if (isStatus) {
            select.onchange = (e) => {
                const stateCodeField = document.querySelectorAll('.enabledInputField[data-fieldname="statecode"]')[0];
                const currentStateCodeValue = stateCodeField.value.split(':')[0].replaceAll(' ', '');
                const stateCodeValueNeeded = e.target.selectedOptions[0].dataset.state;

                if (currentStateCodeValue !== stateCodeValueNeeded) {
                    for (let option of stateCodeField.options) {
                        if (option.value.split(':')[0].replaceAll(' ', '') === stateCodeValueNeeded) {
                            stateCodeField.value = option.value;
                            break;
                        }
                    }
                }
            };
        }
    }

    function createBooleanInput(container, falseOption, trueOption) {
        const value = window.originalResponseCopy[container.dataset.fieldname];

        const select = document.createElement('select');

        let selectHtml = "<option value='null'>null</option>"; // empty option for clearing it

        const falseFormatted = 'false : ' + falseOption?.Label?.UserLocalizedLabel?.Label ?? 'false';
        const trueFormatted = 'true : ' + trueOption?.Label?.UserLocalizedLabel?.Label ?? 'true';
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
        selectTableDiv.append('      Table: ');
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
        lookupEditLinkDiv.classList.add('validLookupEditLinks')

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
        queryDiv.append('      Search text: ');
        queryDiv.append(queryInput);
        editMenuDiv.append(queryDiv);

        const selectFilterField = document.createElement('select');
        selectFilterField.dataset.fieldname = fieldName;
        selectFilterField.style.height = '22px';

        const resultsSelect = document.createElement('select');
        const resultsDiv = document.createElement('div');
        resultsDiv.append('      Search results:');
        resultsDiv.append(resultsSelect);

        // resultsSelect.dataset.id = id;
        // resultsSelect.dataset.originalid = id;
        // resultsSelect.dataset.originalname = name;
        // resultsSelect.dataset.originalNavigationProperty = 'test123';
        // resultsSelect.dataset.originalLookupType = lookupType;

        const selectFilterFieldDefault = document.createElement('option');
        selectFilterFieldDefault.value = '_primary';
        selectFilterFieldDefault.innerText = 'primary column';
        selectFilterField.appendChild(selectFilterFieldDefault);

        editMenuDiv.append(resultsDiv);

        const selectFilterFieldDiv = document.createElement('div');
        selectFilterFieldDiv.append('      Search by: ');
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
        input.dataset.navigationproperty = await retrieveNavigationProperty(logicalName, input.dataset.fieldname);
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

        editMenuDiv.append('      Row id:');
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

    async function deleteRecord(logicalName, pluralName, id) {
        const deleteLink = document.getElementsByClassName('deleteLink')[0];
        deleteLink.style.display = 'none';

        if (confirm("Are you sure you want to delete this row?")) {
            const requestUrl = apiUrl + pluralName + '(' + id + ')';
            const result = await fetch(requestUrl, { method: 'DELETE' });

            if (result.status === 204) {
                if (confirm('Row was deleted. Click OK to close this page.')) {
                    window.close();
                }
            } else {
                const json = await result.json();

                if (json.error) {
                    alert(json.error.message);
                    const deleteLink = document.getElementsByClassName('deleteLink')[0];
                    deleteLink.style.display = 'unset';
                }
            }
        }

        deleteLink.style.display = 'unset';
    }

    async function editRecord(logicalName, pluralName, id, isCreateMode) {
        window.pfwaMode = isCreateMode ? 'create' : 'update';

        const editLink = document.getElementsByClassName('editLink')[0];
        editLink.style.display = 'none';

        await initEntityClientMetadataForCurrentRecord();

        let attributesMetadata = null;
        if (isCreateMode) {
            attributesMetadata = await retrieveCreatableAttributes(logicalName);
            let attributesMetadataForUpdate = await retrieveUpdateableAttributes(logicalName);

            // the statecode is creatable but not listed as such
            // so add 'manually'
            let stateCodeAttribute = attributesMetadataForUpdate.find(a => a.LogicalName === 'statecode');
            if (stateCodeAttribute != null) {
                let stateCodeAttributeAlreadyAdded = attributesMetadata.find(a => a.LogicalName === 'statecode');
                if (stateCodeAttributeAlreadyAdded == null) {
                    attributesMetadata.push(stateCodeAttribute);
                }
            }
        } else {
            attributesMetadata = await retrieveUpdateableAttributes(logicalName);
        }

        const optionSetMetadata = await retrieveOptionSetMetadata(logicalName);
        const multiSelectOptionSetMetadata = await retrieveMultiSelectOptionSetMetadata(logicalName);
        const stateMetadata = await retrieveStateMetadata(logicalName);
        const statusMetadata = await retrieveStatusMetadata(logicalName);

        const booleanMetadata = await retrieveBooleanFieldMetadata(logicalName);

        const inputContainers = document.getElementsByClassName('mainPanel')[0].getElementsByClassName('inputContainer');
        const inputContainersArray = [...inputContainers];

        for (let attribute of attributesMetadata) {
            let container = inputContainersArray.find(c => c.dataset.fieldname === attribute.LogicalName);

            if (container == null) {
                continue;
            }

            if (attribute.LogicalName === 'stageid' || attribute.LogicalName === '_stageid_value') {
                // this field causes many issues for some reason
                // just leave it out
                continue;
            }

            const attributeType = attribute.AttributeType;
            if (attributeType === 'String' || attributeType === 'EntityName') {
                createInput(container, false, 'string');
            }
            else if (attributeType === 'Owner') {
                createLookupInput(container, ['systemuser', 'team']);
            }
            else if (attributeType === 'Memo') {
                createInput(container, true, 'memo');
            } else if (attributeType === 'Picklist') {
                const fieldOptionSetMetadata = optionSetMetadata.find(osv => osv.LogicalName === attribute.LogicalName);
                if (fieldOptionSetMetadata) {
                    const fieldOptionset = fieldOptionSetMetadata.GlobalOptionSet || fieldOptionSetMetadata.OptionSet;
                    createOptionSetValueInput(container, fieldOptionset.Options, true, true, false)
                }
            } else if (attributeType === 'State') {
                const fieldStateMetadata = stateMetadata.find(osv => osv.LogicalName === attribute.LogicalName);
                if (fieldStateMetadata) {
                    const fieldOptionset = fieldStateMetadata.GlobalOptionSet || fieldStateMetadata.OptionSet;
                    createOptionSetValueInput(container, fieldOptionset.Options, false, false, false)
                }
            } else if (attributeType === 'Status') {
                const fieldStatusMetadata = statusMetadata.find(osv => osv.LogicalName === attribute.LogicalName);
                if (fieldStatusMetadata) {
                    const fieldOptionset = fieldStatusMetadata.GlobalOptionSet || fieldStatusMetadata.OptionSet;
                    createOptionSetValueInput(container, fieldOptionset.Options, false, true, true)
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
                createInput(container, false, 'datetime');
            } else if (attributeType === 'Uniqueidentifier') {
                createInput(container, false, 'uid');
            } else if (attributeType === 'Virtual') {
                if (attribute.AttributeTypeName?.Value === 'MultiSelectPicklistType') {
                    const fieldOptionSetMetadata = multiSelectOptionSetMetadata.find(osv => osv.LogicalName === attribute.LogicalName);
                    if (fieldOptionSetMetadata) {
                        const fieldOptionset = fieldOptionSetMetadata.GlobalOptionSet || fieldOptionSetMetadata.OptionSet;
                        createMultiSelectOptionSetValueInput(container, fieldOptionset.Options)
                    }
                }
            }
        }

        const notEnabledContainers = document.getElementsByClassName('mainPanel')[0].getElementsByClassName('containerNotEnabled');

        const length = notEnabledContainers.length;
        for (let i = 0; i < length; i++) {
            notEnabledContainers[0].remove();
        }

        const editMenuDiv = document.getElementsByClassName('mainPanel')[0].getElementsByClassName('editMenuDiv')[0];
        editMenuDiv.style.display = 'inline';

        const cancelLink = document.getElementsByClassName('mainPanel')[0].getElementsByClassName('cancelLink')[0];
        cancelLink.onclick = function () {
            reloadPage(pluralName);
        }

        const submitLink = document.getElementsByClassName('mainPanel')[0].getElementsByClassName('submitLink')[0];
        submitLink.style.display = null;
        submitLink.onclick = async function () {
            cancelLink.style.display = 'none';
            submitLink.style.display = 'none';
            await submitEdit(pluralName, id, isCreateMode);
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

        applySettings();
    }


    function applySettings() {
        try {
            const savedSettings = JSON.parse(localStorage.getItem('pfwaSettings'));
            if (savedSettings) {
                document.getElementById('saveSettingsInput').checked = true;
                const impersonateAnotherUserCheckbox = document.getElementsByClassName('impersonateAnotherUserCheckbox')[0];
                const impersonateAnotherUserField = document.getElementsByClassName('impersonateAnotherUserSelect')[0];
                const impersonateAnotherUserInput = document.getElementsByClassName('impersonateAnotherUserInput')[0];

                if (savedSettings.CallerObjectId) {
                    impersonateAnotherUserField.value = 'azureactivedirectoryobjectid';
                    impersonateAnotherUserCheckbox.click();
                    impersonateAnotherUserInput.value = savedSettings.CallerObjectId;
                    impersonateAnotherUserInput.dispatchEvent(new Event('input', { bubbles: true }));
                } else if (savedSettings.MSCRMCallerID) {
                    impersonateAnotherUserField.value = 'systemuserid';
                    impersonateAnotherUserCheckbox.click();
                    impersonateAnotherUserInput.value = savedSettings.MSCRMCallerID;
                    impersonateAnotherUserInput.dispatchEvent(new Event('input', { bubbles: true }));
                }

                if (savedSettings.bypassCustomPluginExecution) {
                    const bypassCustomPluginExecution = document.getElementsByClassName('bypassPluginExecutionBox')[0];
                    bypassCustomPluginExecution.checked = true;
                }

                if (savedSettings.bypassFlowExecution) {
                    const bypassFlowExecution = document.getElementsByClassName('bypassFlowExecutionBox')[0];
                    bypassFlowExecution.checked = true;
                }

                const previewChangesBeforeSaving = document.getElementsByClassName('previewChangesBeforeSavingBox')[0];

                if (savedSettings.wasPreviewEnabled) {
                    previewChangesBeforeSaving.checked = true;
                } else {
                    previewChangesBeforeSaving.checked = false;
                }
            }
        } catch {
            // ignore
        }
    }

    function resetSubmitControls() {
        const saveInProgressDiv = document.getElementsByClassName('saveInProgressDiv')[0];
        const submitLink = document.getElementsByClassName('mainPanel')[0].getElementsByClassName('submitLink')[0];
        const cancelLink = document.getElementsByClassName('mainPanel')[0].getElementsByClassName('cancelLink')[0];

        saveInProgressDiv.style.display = 'none';
        cancelLink.style.display = null;
        submitLink.style.display = null;

        destroyPreview();
        enableAllInputs();
    }

    async function submitEdit(pluralName, id, isCreateMode) {
        const previewChangesBeforeSaving = document.getElementsByClassName('mainPanel')[0].getElementsByClassName('previewChangesBeforeSavingBox')[0].checked;

        const changedFields = {};
        const enabledFields = document.getElementsByClassName('mainPanel')[0].getElementsByClassName('enabledInputField');
        for (let input of enabledFields) {
            let originalValue = window.originalResponseCopy[input.dataset.fieldname];
            const dataType = input.dataset.datatype;
            const inputValue = input.value;
            const fieldName = input.dataset.fieldname;

            let value = '';
            if (dataType === 'string' || dataType === 'memo' || dataType === 'uid') {
                if (inputValue === '' && input.dataset.isEmptyString !== 'true') {
                    value = null;
                } else {
                    value = inputValue;
                }
            }
            else if (dataType === 'option') {
                if (!inputValue) {
                    // the select needs to contain a value always, if not, an error happened
                    alert('there was an error parsing the field ' + fieldName);
                    resetSubmitControls();
                    return;
                }
                if (inputValue === 'null') {
                    value = null;
                } else {
                    value = parseInt(inputValue.split(':')[0].replace(' ', ''));
                }
            } else if (dataType === 'multiselectoption') {
                input.querySelectorAll('input').forEach(input => {
                    if (input.checked) {
                        if (value !== '') {
                            value += ',';
                        }

                        value += input.dataset.value;
                    }
                })
                // normalize to null if no options checked
                if (value === '') {
                    value = null;
                }
            } else if (dataType === 'int') {
                if (inputValue == null || inputValue === undefined || inputValue === '') {
                    value = null;
                } else {
                    value = parseInt(inputValue);
                    if (isNaN(value)) {
                        alert(fieldName + ' is a whole number. The value ' + inputValue + ' is not compatible.');
                        resetSubmitControls();
                        return;
                    }

                    if (/^-?\d+$/.test(inputValue) === false) {
                        alert(fieldName + ' is a whole number. The value ' + inputValue + ' is not compatible.');
                        resetSubmitControls();
                        return;
                    }
                }
            } else if (dataType === 'decimal' || dataType == 'money' || dataType == 'float') {
                if (inputValue == null || inputValue === undefined || inputValue === '') {
                    value = null;
                } else {
                    if (inputValue.includes(',')) {
                        alert(`${fieldName} is a column of type '${dataType}' and contains a comma (,). Use a dot (.) as the separator.`);
                        resetSubmitControls();
                        return;
                    }

                    value = parseFloat(inputValue);
                    if (isNaN(value)) {
                        alert(`${fieldName} is a column of type '${dataType}'. The value ${inputValue} is not compatible.`);
                        resetSubmitControls();
                        return;
                    }

                    if (/^-?[0-9]\d*(\.\d+)?$/.test(inputValue) === false) {
                        alert(`${fieldName} is a column of type '${dataType}'. The value ${inputValue} is not compatible.`);
                        resetSubmitControls();
                        return;
                    }
                }
            } else if (dataType === 'bool') {
                if (!inputValue) {
                    // the select needs to contain a value always, if not, an error happened
                    alert('there was an error parsing the field ' + fieldName);
                    resetSubmitControls();
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
                        resetSubmitControls();
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
                    resetSubmitControls();
                    return;
                }

                const pluralName = await retrievePluralName(tableSelectValue)

                value = `/${pluralName}(${inputValue})`;

                if (input.dataset.editmode === 'update') {
                    if (inputValue == null || inputValue === '') {
                        alert('Error for lookup field: ' + fieldName + '. The field was marked for update but it is empty. If you do not want to edit this field, hit "undo changes".');
                        resetSubmitControls();
                        return;
                    }

                    if (inputValue.length !== 36) {
                        alert('Error for lookup field: ' + fieldName + '. The value ' + value + ' is not a valid guid.');
                        resetSubmitControls();
                        return;
                    }
                } else if (input.dataset.editmode === 'makenull') {
                    value = null;
                } else {
                    // skip if not action is required
                    continue;
                }
            } else if (dataType === 'datetime') {
                if (inputValue === '') {
                    value = null;
                } else {
                    if (Date.parse(value) === NaN) {
                        alert('Error for datetime field: ' + fieldName + '. The value ' + value + ' is not a valid datetime.');
                        resetSubmitControls();
                        return;
                    }
                    value = inputValue;
                }
            }

            if (value !== originalValue && !(value === '' && originalValue == null)) {
                if (dataType === 'memo') {
                    if (originalValue?.replaceAll('\r\n', '\n') !== value) {
                        changedFields[fieldName] = value;
                    }
                } else if (dataType == 'lookup') {
                    const lookupFieldName = `${input.dataset.navigationproperty}@odata.bind`;

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
                resetSubmitControls();
                return;
            }

            if (impersonateAnotherUserInput?.length !== 36) {
                alert('User impersonation input error: ' + impersonateAnotherUserInput + ' is not a valid guid.');
                resetSubmitControls();
                return;
            }

            if (impersonateAnotherUserField == 'systemuserid') {
                impersonateHeader['MSCRMCallerID'] = impersonateAnotherUserInput;
            } else if (impersonateAnotherUserField == 'azureactivedirectoryobjectid') {
                impersonateHeader['CallerObjectId'] = impersonateAnotherUserInput;
            } else {
                alert('This should not happen. Wrong value in impersonateAnotherUserSelect: ' + impersonateAnotherUserField);
                resetSubmitControls();
                return;
            }
        }

        if (!!previewChangesBeforeSaving) {
            previewChanges(changedFields, pluralName, id, impersonateHeader, isCreateMode);
        } else {
            await commitSave(pluralName, id, changedFields, impersonateHeader, isCreateMode, false);
        }
    }

    async function commitSave(pluralName, id, changedFields, impersonateHeader, isCreateMode, wasPreviewEnabled) {
        let requestUrl = '';
        let method = '';

        if (isCreateMode) {
            requestUrl = apiUrl + pluralName;
            method = 'POST';
        } else {
            requestUrl = apiUrl + pluralName + '(' + id + ')';
            method = 'PATCH';
        }

        Object.keys(changedFields)
            .filter(key => key.includes('____lookupOverride'))
            .forEach(key => delete changedFields[key]);

        let headers = {
            'accept': 'application/json',
            'content-type': 'application/json',
            'OData-MaxVersion': '4.0',
            'OData-Version': '4.0',
        };

        if (!isCreateMode) {
            headers['If-Match'] = '*'
        } else {
            headers['Prefer'] = 'return=representation'
        }

        headers = { ...headers, ...impersonateHeader }

        const bypassCustomPluginExecution = document.getElementsByClassName('bypassPluginExecutionBox')[0].checked;
        if (!!bypassCustomPluginExecution) {
            headers['MSCRM.BypassCustomPluginExecution'] = true;
        }

        const bypassFlowExecution = document.getElementsByClassName('bypassFlowExecutionBox')[0].checked;
        if (!!bypassFlowExecution) {
            headers['MSCRM.SuppressCallbackRegistrationExpanderJob'] = true;
        }

        saveSettings(impersonateHeader, bypassCustomPluginExecution, bypassFlowExecution, wasPreviewEnabled);

        const saveInProgressDiv = document.getElementsByClassName('saveInProgressDiv')[0];
        saveInProgressDiv.style.display = null;

        const response = await fetch(requestUrl, {
            method: method,
            headers: headers,
            body: JSON.stringify(changedFields)
        });

        if (response.ok) {
            if (isCreateMode) {
                const json = await response.json();

                const primaryIdAttribute = (await retrieveLogicalAndPrimaryKeyAndPrimaryName(pluralName)).primaryIdAttribute;
                window.location.href = apiUrl + pluralName + '(' + json[primaryIdAttribute] + ')#p';
            } else {
                reloadPage(pluralName);
            }
        } else {
            const errorText = await response.text();
            console.error(`${response.status} - ${errorText}`);
            window.alert(`${response.status} - ${errorText}`);

            resetSubmitControls();
        }
    }

    function saveSettings(impersonateHeader, bypassCustomPluginExecution, bypassFlowExecution, wasPreviewEnabled) {
        try {
            const saveSettings = document.getElementById('saveSettingsInput').checked;
            if (!!saveSettings) {

                let settings = impersonateHeader || {};
                if (!!bypassCustomPluginExecution) {
                    settings.bypassCustomPluginExecution = true;
                }

                if (!!bypassFlowExecution) {
                    settings.bypassFlowExecution = true;
                }

                settings.wasPreviewEnabled = wasPreviewEnabled;

                localStorage.setItem('pfwaSettings', JSON.stringify(settings));
            } else {
                localStorage.removeItem('pfwaSettings');
            }
        } catch {
            // ignore
        }
    }

    async function reloadPage(pluralName) {
        if (pluralName === 'workflows') {
            window.location.reload(); // the monaco editor fails us when re-initializing, so just reload
        } else {
            await makeItPretty();
        }
    }

    function stringContains(str, value) {
        return str.indexOf(value) !== -1;
    }

    function isAnnotation(key) {
        return stringContains(key, formattedValueType) || stringContains(key, navigationPropertyType) || stringContains(key, lookupType);
    }


    function createMonacoEditorControls(mainPanel, recordId) {
        if (checkIfJsonViewerEnabled()) {
            const btn = document.createElement('button');
            btn.style = `
                height: 30px;
                width: auto;
                margin-right: 160px;
                margin-top: 10px;
                position: absolute;
                right: 10px;
                padding:0;
                font-size:24;
                padding: 0px 4px;
                disabled;
                `
            btn.innerHTML = '<div>Currently, editing Flows is only possible in Chrome. Work in progress.</div>';
            btn.setAttribute('disabled', 'true');

            mainPanel.prepend(btn);
            return;
        }

        mainPanel.dataset.chromeRuntimeUrl = chrome.runtime.getURL('');
        mainPanel.dataset.flowName = originalResponseCopy.name;
        mainPanel.dataset.recordId = recordId;
        mainPanel.dataset.apiUrl = apiUrl;

        let scriptTag = document.createElement('script');
        scriptTag.src = chrome.runtime.getURL('libs/monaco/loader.js');
        scriptTag.type = "text/javascript";
        document.head.appendChild(scriptTag);

        window.setTimeout(() => {
            let scriptTagInit = document.createElement('script');
            scriptTagInit.src = chrome.runtime.getURL('libs/monaco/initMonaco.js');
            scriptTagInit.type = "text/javascript";
            document.head.appendChild(scriptTagInit);
        }, 200);
    }

    function customJsonStringify(jsonObj, currentDepth) {
        var padding = "&nbsp;".repeat(currentDepth * 3);
        var nestedPadding = "&nbsp;".repeat(3);

        let json = '';
        if (currentDepth === 0) {
            json = `{`;
        }
        for (const key in jsonObj) {
            let value = jsonObj[key];

            let insertKeyname = '';
            if (key !== '0' && !parseInt(key)) {
                insertKeyname = `${key}: `;
            }
            if (Array.isArray(value)) {
                json += `\n${padding}${nestedPadding}${insertKeyname}[${customJsonStringify(value, currentDepth + 1)}\n${padding}${nestedPadding}]`;
            } else if (typeof value === 'object') {
                json += `\n${padding}${nestedPadding}${insertKeyname}{${customJsonStringify(value, currentDepth + 1)}\n${padding}${nestedPadding}}`;
            } else {
                value = value?.replaceAll('\r\n', '\\n').replaceAll('\n', '\\n');
                json += `\n${padding}${nestedPadding}${insertKeyname}${value}`;
            }
        }

        if (currentDepth === 0) {
            json += `\n}`;
        }
        return json;
    }

    async function prettifyWebApi(jsonObj, htmlElement, pluralName, isPreview, isCreateMode) {
        window.pfwaMode = isCreateMode ? 'create' : 'read';
        if (window.originalResponseCopy == null) {
            window.originalResponseCopy = JSON.parse(JSON.stringify(jsonObj));
        }

        const isMultiple = (jsonObj.value && Array.isArray(jsonObj.value));
        const isSingleColumnValueOnly = Object.keys(jsonObj).length === 2 && jsonObj['@odata.context'] && jsonObj.value;

        let result = await retrieveLogicalAndPrimaryKeyAndPrimaryName(pluralName);

        if (window.location.hash === '#pf' && jsonObj.value.length === 1) {
            const recordId = jsonObj.value[0][result.primaryIdAttribute];
            const newUrl = apiUrl + pluralName + "(" + recordId + ")#p";
            window.location.href = newUrl;
            return;
        }

        let singleRecordId = '';
        let relationShipDefinition = null;

        // the second part of this statement is a bit weird
        // but we do this to check if the context is 'manually' expanded on a relationship
        const inferredRelationship = !isSingleColumnValueOnly && location.pathname.includes(')/');
        if (location.hash === '#pr' || inferredRelationship) {
            let relationShipName = '';
            try {
                relationShipName = location.pathname.split('/').slice(-1)[0];

                relationShipDefinition = await retrieveRelationshipDefinition(relationShipName);

                if (relationShipDefinition == null) {
                    throw `Relationship not found`;
                } else {
                    let overridenPluralName = null;

                    if (relationShipDefinition.RelationshipType === 1) {
                        if (relationShipDefinition.Entity1LogicalName !== result.logicalName && relationShipDefinition.Entity2LogicalName === result.logicalName) {
                            overridenPluralName = await retrievePluralName(relationShipDefinition.Entity1LogicalName);
                        } else if (relationShipDefinition.Entity1LogicalName === result.logicalName && relationShipDefinition.Entity2LogicalName !== result.logicalName) {
                            overridenPluralName = await retrievePluralName(relationShipDefinition.Entity2LogicalName);
                        } else if (relationShipDefinition.Entity1LogicalName === result.logicalName && relationShipDefinition.Entity2LogicalName === result.logicalName) {
                            // self referencing entity. No need to override the context
                        } else {
                            throw `Cannot map relationship to either ${relationShipDefinition.Entity1LogicalName} or ${relationShipDefinition.Entity2LogicalName}`;
                        }
                    } else if (relationShipDefinition.RelationshipType === 0) {
                        if (relationShipDefinition.ReferencingEntity !== result.logicalName && relationShipDefinition.ReferencedEntity === result.logicalName) {
                            overridenPluralName = await retrievePluralName(relationShipDefinition.ReferencingEntity);
                        } else if (relationShipDefinition.ReferencingEntity === result.logicalName && relationShipDefinition.ReferencedEntity !== result.logicalName) {
                            overridenPluralName = await retrievePluralName(relationShipDefinition.ReferencedEntity);
                            window.originalEntityPluralName = pluralName;
                            window.isManyToOneContext = true;
                        } else if (relationShipDefinition.ReferencingEntity === result.logicalName && relationShipDefinition.ReferencedEntity === result.logicalName) {
                            // self referencing entity. No need to override the context
                        } else {
                            throw `Cannot map relationship to either ${relationShipDefinition.ReferencingEntity} or ${relationShipDefinition.ReferencedEntity}`;
                        }
                    } else {
                        throw `Cannot handle RelationshipType ${relationShipDefinition.RelationshipType}`;
                    }

                    if (overridenPluralName != null && overridenPluralName !== '') {
                        // override the context to be the referenced table
                        result = await retrieveLogicalAndPrimaryKeyAndPrimaryName(overridenPluralName);
                        pluralName = overridenPluralName;
                        window.currentEntityPluralName = overridenPluralName;
                    }
                }
            } catch (err) {
                relationShipDefinition = null;
                if (!inferredRelationship) {
                    alert(`Something went wrong with retrieving the relationship definition for '${relationShipName}'. Error message: ${err}`);
                }
            }
        }

        // sdk messages other than retrieve or retrievemultiple
        if (!result.logicalName) {
            jsonObj = await enrichObjectWithHtml(jsonObj, null, null, null, !isPreview, false, 1, null, isCreateMode, null, false);
        } else if (isMultiple) {
            delete jsonObj[etagKey];
            delete jsonObj[metaDataKey];
            delete jsonObj[contextKey];
            delete jsonObj[totalRecordCountKey];
            delete jsonObj[totalRecordCountLimitExceededKey];

            if (!titleSet) {
                if (relationShipDefinition != null) {
                    document.title = 'related ' + pluralName;
                } else {
                    document.title = pluralName;
                }
                titleSet = true;
            }
            const valueKeyWithCount = 'value (' + jsonObj.value.length + ' rows)';

            jsonObj[valueKeyWithCount] = jsonObj.value;
            delete jsonObj.value;

            for (const key in jsonObj[valueKeyWithCount]) {
                jsonObj[valueKeyWithCount][key] = await enrichObjectWithHtml(jsonObj[valueKeyWithCount][key], result.logicalName, pluralName, result.primaryIdAttribute, false, false, 2, result.primaryNameAttribute, isCreateMode, relationShipDefinition, false);
            }
        } else {
            if (!titleSet) {
                document.title = result.logicalName;
                titleSet = true;
            }

            singleRecordId = jsonObj[result.primaryIdAttribute];

            jsonObj = await enrichObjectWithHtml(jsonObj, result.logicalName, pluralName, result.primaryIdAttribute, !isPreview, false, 1, result.primaryNameAttribute, isCreateMode, relationShipDefinition, isSingleColumnValueOnly);
        }

        let json = customJsonStringify(jsonObj, 0);
        json = json.replaceAll('"', '').replaceAll(replacedQuote, escapeHtml('"'));
        json = json.replaceAll(',', '').replaceAll(replacedComma, ',');

        htmlElement.innerText = '';

        const pre = document.createElement('pre');
        pre.style.margin = '0px';
        htmlElement.appendChild(pre).innerHTML = json;

        if (!isPreview) {
            pre.classList.add('mainPanel');
            pre.classList.add('panel');

            pre.style.position = 'relative';

            if (!isCreateMode) {
                const btn = document.createElement('button');
                btn.style = `
                height: 30px;
                width: 122px;
                margin-right: 18px;
                margin-top: 10px;
                position: absolute;
                right: 10px;
                cursor: pointer;
                padding:0;
                font-size:24;
                padding: 0px 4px 0px 4px;
                `

                btn.innerHTML = '<div>View as raw JSON</div>';
                btn.onclick = () => { window.location.href = window.location.href.split('#')[0]; };

                pre.prepend(btn);
            }
        }

        if (!isPreview && !isMultiple && pluralName === 'workflows' && window.originalResponseCopy.hasOwnProperty('clientdata') && window.originalResponseCopy.clientdata?.startsWith('{"')) {
            createMonacoEditorControls(pre, singleRecordId);
            pre.dataset.clientdata = window.originalResponseCopy.clientdata;
        }

        if (!isSingleColumnValueOnly) {
            setPreviewLinkClickHandlers();
            setEditLinkClickHandlers();
            setDeleteLinkClickHandlers();
            setCopyToClipboardHandlers();
            setLookupEditHandlers();
            setDisassociateClickHandlers();
        }

        if (relationShipDefinition == null && result.logicalName && !isCreateMode && !isPreview && pluralName !== 'workflows') {
            setCreateNewRecordButton(pre, result.logicalName);
            // setFilterDataTypesButton(pre);
        }

        if (!isSingleColumnValueOnly && relationShipDefinition == null && result.logicalName && !isCreateMode && !isPreview && pluralName !== 'workflows' && !isMultiple) {
            setBrowseRelationShipsButton(pre, result.logicalName);
        }

        if (!isSingleColumnValueOnly && !isMultiple && !isPreview && result.logicalName != null) {
            setImpersonateUserHandlers();
        }
    }

    function setCreateNewRecordButton(pre, logicalName) {
        const btn = document.createElement('button');
        btn.style = `
            height: 30px;
            width: 122px;
            margin-right: 158px;
            margin-top: 10px;
            position: absolute;
            right: 10px;
            cursor: pointer;
            padding:0;
            font-size:24;
            padding: 0px 4px;
            `

        btn.innerHTML = '<div>Create new row</div>';
        btn.onclick = async () => {
            btn.style.display = 'none'
            await handleCreateNewRecord(logicalName)
        };

        pre.prepend(btn);
    }

    function setFilterDataTypesButton(pre) {
        const btn = document.createElement('button');
        btn.style = `
            height: 30px;
            width: 122px;
            margin-right: 298px;
            margin-top: 10px;
            position: absolute;
            right: 10px;
            cursor: pointer;
            padding:0;
            font-size:24;
            padding: 0px 4px;
            `

        btn.innerHTML = '<div>Filter on data type</div>';
        btn.onclick = async () => {
            handleFilterOnDataType(pre);
        };

        pre.prepend(btn);
    }

    function handleFilterOnDataType(pre) {
        const flyOut = document.createElement('div');
        flyOut.id = 'dataTypeFilterFlyOut';
        flyOut.style = `
            height: 194px;
            width: 99px;
            margin-right: 298px;
            margin-top: 50px;
            position: absolute;
            right: 10px;
            border-style: solid;
            border-width: 2px;
            padding: 10px;
            z-index:100;
        `
        pre.prepend(flyOut);

        createDataFilterControl('Boolean', 'Boolean', flyOut);
        createDataFilterControl('Choice', 'Picklist|State|Status|Virtual', flyOut);
        createDataFilterControl('Date', 'DateTime', flyOut);
        createDataFilterControl('Guid', 'Uniqueidentifier', flyOut);
        createDataFilterControl('Lookup', 'Lookup|Owner', flyOut);
        createDataFilterControl('Number', 'Integer|Decimal|Double|Money|BigInt', flyOut);
        createDataFilterControl('String', 'String|Memo', flyOut);
        createDataFilterControl('Hide null', 'hideNull', flyOut);

        flyOut.onmouseleave = () => flyOut.remove();
    }

    function createDataFilterControl(label, dataType, parent) {
        const div = document.createElement('div');
        div.style = `
            margin: 0 0 10px 0;
            align-content: center;
            display: flex;
            cursor: pointer;
        `
        const checkbox = document.createElement('input');
        checkbox.classList.add('dataFilterCheckBox')
        checkbox.style = `
            width: auto;
            margin: 0 5px 0 0;
            cursor: pointer;
        `
        checkbox.type = 'checkbox';

        if (dataTypeFilters.filter(str => str === dataType).length > 0) {
            checkbox.checked = true;
        }

        checkbox.onclick = async (e) => {
            e.stopPropagation()

            if (checkbox.checked) {
                dataTypeFilters.push(dataType);
            } else {
                dataTypeFilters = dataTypeFilters.filter(str => str !== dataType);
            }

            if (dataTypeFilters.length > 0) {

                const allKeySpans = document.querySelectorAll('.keySpan');

                for (let keySpan of allKeySpans) {
                    const key = keySpan.dataset.key;
                    const keyForDataTypeFilter = key.split('@')[0];
                    const attributeType = keySpan.dataset.attributetype;
                    const cls = keySpan.dataset.cls;

                    const hideNulls = dataTypeFilters.filter(str => str === 'hideNull').length > 0;
                    if (hideNulls) {
                        if (cls === 'null') {

                            keySpan.style.display = 'none';
                            keySpan.nextSibling.nextSibling.style.display = 'none';
                            continue;
                        }
                    }

                }
                //     const keyForDataTypeFilter = key.split('@')[0];
                //     const attributesWithTypes = await retrieveAttributesWithTypes(logicalName);

                //     const attribute = attributesWithTypes.filter(a => a.LogicalName === keyForDataTypeFilter || '_' + a.LogicalName + '_value' === keyForDataTypeFilter)[0];

                //     if (attribute == null) {
                //         alert(`Could not map attribute ${keyForDataTypeFilter} for filtering on data type.`);
                //         delete ordered[keyForDataTypeFilter];
                //         delete ordered[keyForDataTypeFilter + formattedValueType];
                //         delete ordered[keyForDataTypeFilter + navigationPropertyType];
                //         delete ordered[keyForDataTypeFilter + lookupType];
                //         continue;
                //     }

                //     const hideNulls = dataTypeFilters.filter(str => str === 'hideNull').length > 0;
                //     if (hideNulls) {
                //         if (cls === 'null') {
                //             delete ordered[keyForDataTypeFilter];
                //             delete ordered[keyForDataTypeFilter + formattedValueType];
                //             delete ordered[keyForDataTypeFilter + navigationPropertyType];
                //             delete ordered[keyForDataTypeFilter + lookupType];
                //             continue;
                //         }
                //     }

                //     const showOnlySelectedAny = dataTypeFilters.filter(str => str !== 'hideNull').length > 0;

                //     if (showOnlySelectedAny) {
                //         const shouldHideThisType = dataTypeFilters.filter(str => str.includes(attribute.AttributeType)).length === 0;
                //         if (shouldHideThisType) {
                //             delete ordered[keyForDataTypeFilter];
                //             delete ordered[keyForDataTypeFilter + formattedValueType];
                //             delete ordered[keyForDataTypeFilter + navigationPropertyType];
                //             delete ordered[keyForDataTypeFilter + lookupType];
                //             continue;
                //         }
                //     }
            }
        };

        div.onclick = async (e) => {
            checkbox.click();
        };

        div.appendChild(checkbox);
        div.append(label);
        parent.appendChild(div);
    }

    function setBrowseRelationShipsButton(pre, logicalName) {
        const btn = document.createElement('button');
        btn.style = `
            height: 30px;
            width: 122px;
            margin-right: 298px;
            margin-top: 10px;
            position: absolute;
            right: 10px;
            cursor: pointer;
            padding:0;
            font-size:24;
            padding: 0px 4px;
            `

        btn.innerHTML = '<div>View relationships</div>';
        btn.onclick = async () => {
            btn.style.display = 'none';
            await handleBrowseRelationships(logicalName)
        };

        pre.prepend(btn);
    }

    async function handleBrowseRelationships(logicalName) {
        const relationships = await retrieveRelationships();
        if (relationships == null) {
            alert(`An error occurred while retrieving relationships. Cannot continue.`);
            return;
        }

        document.body.innerText = '';

        const relationshipBrowser = document.createElement('pre');
        relationshipBrowser.id = 'relationshipBrowser';
        relationshipBrowser.style.display = 'grid';
        document.body.appendChild(relationshipBrowser);

        const closeButton = document.createElement('button');
        relationshipBrowser.appendChild(closeButton);
        closeButton.style.position = 'fixed';
        closeButton.style.height = '30px';
        closeButton.style.top = '20px';
        closeButton.style.right = '20px';
        closeButton.innerText = 'Close relationships';
        closeButton.onclick = () => window.location.reload();

        const filterLabel = document.createElement('h3');
        filterLabel.innerText = 'Filter relationships';
        relationshipBrowser.appendChild(filterLabel);

        const filterInput = document.createElement('input');
        filterInput.type = 'text';
        filterInput.style.maxWidth = '500px';
        relationshipBrowser.appendChild(filterInput);

        const relationshipsDiv = document.createElement('div');
        relationshipBrowser.appendChild(relationshipsDiv);

        filterInput.oninput = () => {
            relationshipsDiv.innerText = '';
            updateRelationshipBrowserContents(relationshipsDiv, relationships, filterInput.value, logicalName);
        };

        updateRelationshipBrowserContents(relationshipsDiv, relationships, null, logicalName);
    }

    function updateRelationshipBrowserContents(container, relationships, optionalFilter, logicalName) {
        let manyToMany = relationships.ManyToManyRelationships
        let oneToMany = relationships.OneToManyRelationships;
        let manyToOne = relationships.ManyToOneRelationships;

        if (optionalFilter != null && optionalFilter !== '') {
            optionalFilter = optionalFilter.toLowerCase();

            manyToMany = manyToMany.filter(r =>
                r.Entity1LogicalName?.toLowerCase()?.includes(optionalFilter) ||
                r.Entity2LogicalName?.toLowerCase()?.includes(optionalFilter) ||
                r.IntersectEntityName?.toLowerCase()?.includes(optionalFilter) ||
                r.Entity1IntersectAttribute?.toLowerCase()?.includes(optionalFilter) ||
                r.Entity2IntersectAttribute?.toLowerCase()?.includes(optionalFilter) ||
                r.Entity1NavigationPropertyName?.toLowerCase()?.includes(optionalFilter) ||
                r.Entity2NavigationPropertyName?.toLowerCase()?.includes(optionalFilter) ||
                r.SchemaName?.toLowerCase()?.includes(optionalFilter)
            );

            oneToMany = oneToMany.filter(r =>
                r.ReferencedAttribute?.toLowerCase()?.includes(optionalFilter) ||
                r.ReferencedEntity?.toLowerCase()?.includes(optionalFilter) ||
                r.ReferencingAttribute?.toLowerCase()?.includes(optionalFilter) ||
                r.ReferencingEntity?.toLowerCase()?.includes(optionalFilter) ||
                r.ReferencedEntityNavigationPropertyName?.toLowerCase()?.includes(optionalFilter) ||
                r.ReferencingEntityNavigationPropertyName?.toLowerCase()?.includes(optionalFilter) ||
                r.SchemaName?.toLowerCase()?.includes(optionalFilter)
            );

            manyToOne = manyToOne.filter(r =>
                r.ReferencedAttribute?.toLowerCase()?.includes(optionalFilter) ||
                r.ReferencedEntity?.toLowerCase()?.includes(optionalFilter) ||
                r.ReferencingAttribute?.toLowerCase()?.includes(optionalFilter) ||
                r.ReferencingEntity?.toLowerCase()?.includes(optionalFilter) ||
                r.ReferencedEntityNavigationPropertyName?.toLowerCase()?.includes(optionalFilter) ||
                r.ReferencingEntityNavigationPropertyName?.toLowerCase()?.includes(optionalFilter) ||
                r.SchemaName?.toLowerCase()?.includes(optionalFilter)
            );
        }

        const manyToManyHeader = document.createElement('h2');
        manyToManyHeader.innerText = 'Many to Many Relationships';
        container.appendChild(manyToManyHeader);

        for (let relationship of manyToMany) {
            createRelationshipElementForManyToMany(relationship, logicalName, container);
        }

        const manyToOneHeader = document.createElement('h2');
        manyToOneHeader.innerText = 'Many to One Relationships (lookups)';
        container.appendChild(manyToOneHeader);

        for (let relationship of manyToOne) {
            createRelationshipElementForManyToOne(relationship, logicalName, container);
        }

        const oneToManyHeader = document.createElement('h2');
        oneToManyHeader.innerText = 'One to Many Relationships';
        container.appendChild(oneToManyHeader);

        for (let relationship of oneToMany) {
            createRelationshipElementForOneToMany(relationship, logicalName, container);
        }

        container.style.display = 'grid';
    }

    function createRelationshipElementForManyToMany(relationship, logicalName, container) {
        let navigationProperty = null;

        if (relationship.Entity1LogicalName === logicalName) {
            navigationProperty = relationship.Entity1NavigationPropertyName;
        } else if (relationship.Entity2LogicalName === logicalName) {
            navigationProperty = relationship.Entity2NavigationPropertyName;
        } else {
            alert(`Could find the navigation property for many to many relationship ${relationship.SchemaName}`);
            return;
        }

        const link = createRelationshipLink(navigationProperty);
        container.appendChild(link);
    }

    function createRelationshipElementForManyToOne(relationship, logicalName, container) {
        let navigationProperty = null;

        if (relationship.ReferencingEntity === logicalName) {
            navigationProperty = relationship.ReferencingEntityNavigationPropertyName;
        } else {
            alert(`Could find the navigation property for many to one relationship '${relationship.SchemaName}'`);
            return;
        }

        const link = createRelationshipLink(navigationProperty);

        if (window.originalResponseCopy['_' + navigationProperty + '_value'] == null && foundNavigationProperties.find(n => n == navigationProperty) == null) {
            link.href = 'javascript:'
            link.style['text-decoration'] = 'none';
            link.style['pointer-events'] = 'none';
        }

        container.appendChild(link);
    }

    function createRelationshipElementForOneToMany(relationship, logicalName, container) {
        let navigationProperty = null;

        if (relationship.ReferencedEntity === logicalName) {
            navigationProperty = relationship.ReferencedEntityNavigationPropertyName;
        } else {
            alert(`Could find the navigation property for one to many relationship '${relationship.SchemaName}'`);
            return;
        }

        const link = createRelationshipLink(navigationProperty);
        container.appendChild(link);
    }

    function createRelationshipLink(navigationProperty) {
        const element = document.createElement('a');
        element.target = '_blank';
        element.style.margin = '4px 0px';
        element.style.width = 'fit-content';
        element.href = window.location.pathname + '/' + navigationProperty + '#pr';
        element.innerText = navigationProperty;
        return element;
    }

    async function handleCreateNewRecord(logicalName) {
        const creatableAttributes = await retrieveCreatableAttributes(logicalName);
        let jsonObject = {};

        for (let attribute of creatableAttributes) {
            if (attribute.AttributeType === 'Lookup' && attribute.Targets.length > 0) {
                jsonObject["_" + attribute.LogicalName + "_value"] = null;
            } else if (attribute.AttributeType === 'Owner') {
                jsonObject["_" + attribute.LogicalName + "_value"] = null;
            }
            else {
                jsonObject[attribute.LogicalName] = null;
            }
        }

        // 'manually' add this in, because we want to be able to see this
        // it's not listed as creatable, but it actually is
        jsonObject['statecode'] = null;

        await prettifyWebApi(jsonObject, document.body, window.currentEntityPluralName, false, true);
        await editRecord(logicalName, window.currentEntityPluralName, null, true);
    }

    function previewChanges(changedFields, pluralName, id, impersonateHeader, isCreateMode) {
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
            if (!isCreateMode) {
                change.old = originalValue;
            }
            change.new = updatedValue;
            changes.push(change);
        }

        const table = tableFromChanges(changes, isCreateMode);

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

        undoAllLink.onclick = () => reloadPage(pluralName);

        editMenu.appendChild(undoAllLink);

        const lineBreak2 = document.createElement('br');
        editMenu.appendChild(lineBreak2);
        editMenu.append('    ');

        const submitChangesLink = document.createElement('a');
        submitChangesLink.innerText = 'Commit Save';
        submitChangesLink.href = 'javascript:';

        // create this callback so we enclose the values we need when saving
        const saveCallback = async function () {
            submitChangesLink.style.display = 'none';
            undoAllLink.style.display = 'none';
            await commitSave(pluralName, id, changedFields, impersonateHeader, isCreateMode, true);
        }

        submitChangesLink.onclick = saveCallback;

        editMenu.appendChild(submitChangesLink);
    }

    function enableAllInputs() {
        const inputs = document.getElementsByTagName('input');
        for (let i = 0; i < inputs.length; i++) {
            inputs[i].disabled = false;
        }
        const selects = document.getElementsByTagName('select');
        for (let i = 0; i < selects.length; i++) {
            if (selects[i].dataset.fieldname !== 'statecode') {
                selects[i].disabled = false;
            }
        }
        const textareas = document.getElementsByTagName('textarea');
        for (let i = 0; i < textareas.length; i++) {
            textareas[i].disabled = false;
        }
        const lookupEditLinks = document.getElementsByClassName('validLookupEditLinks');
        for (let i = 0; i < lookupEditLinks.length; i++) {
            lookupEditLinks[i].style.display = 'unset';
        }
    }

    function destroyPreview() {
        const editMenu = document.getElementById('previewChangesDiv');
        editMenu.innerHTML = '  ';
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
        const lookupEditLinks = document.getElementsByClassName('validLookupEditLinks');
        for (let i = 0; i < lookupEditLinks.length; i++) {
            lookupEditLinks[i].style.display = 'none';
        }
    }

    async function previewRecord(pluralName, url) {
        document.getElementsByClassName('mainPanel')[0].classList.add("prePreviewed");
        document.body.classList.add("bodyPreviewed");

        const newDiv = document.createElement('div');
        newDiv.classList.add('panel');
        newDiv.classList.add('previewPanel');
        newDiv.classList.add('prePreviewed');

        document.body.appendChild(newDiv);

        newDiv.style = 'position:relative;'

        const response = await odataFetch(url, false);

        await prettifyWebApi(response, newDiv, pluralName, true, false);

        const btn = document.createElement('button');
        btn.style = `
                height: 30px;
                width: 122px;
                margin-right: 18px;
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
                document.getElementsByClassName('mainPanel')[0].classList.remove("prePreviewed");
                document.body.classList.remove("bodyPreviewed");
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

    function checkIfJsonViewerEnabled() {
        try {
            if (navigator.userAgent.match(/edg/i)) {
                return document.body.getAttribute('data-code-mirror') != null || document.getElementById('settings_button') != null || document.getElementById('code_folding') != null;
            }
        } catch {
            // do nothing
        }
        return false;
    }

    async function makeItPretty(optionalJsobObject) {
        if (window.location.hash === '#pf') {
            document.body.innerText = 'Loading your flow...';
        }
        let jsonObject = {};
        if (optionalJsobObject) {
            jsonObject = optionalJsobObject;
        } else {
            jsonObject = await odataFetch(window.location.href, false);
        }

        window.currentEntityPluralName = window.location.pathname.split('(').shift().split('/').pop();

        await prettifyWebApi(jsonObject, document.body, window.currentEntityPluralName, false);

        if (window.location.hash === '#pf') {
            return;
        }

        addMainCss();

        const transParentOverlay = document.createElement('div');
        transParentOverlay.id = 'transParentOverlay';
        transParentOverlay.style.display = 'none';
        transParentOverlay.onclick = () => {
            const multiSelectDivs = document.querySelectorAll('.multiSelectDiv');
            multiSelectDivs.forEach(d => d.style.display = 'none');
            transParentOverlay.style.display = 'none';
        };

        document.body.appendChild(transParentOverlay);

        // do this for Edge, who overrides our styles
        for (var panel of document.getElementsByClassName('panel')) {
            panel.style.display = 'block';
        }
    }

    function clearCss() {
        const head = document.getElementsByTagName('head')[0];
        const styles = head.getElementsByTagName('style');
        for (let style of styles) {
            head.removeChild(style);
        }
    }

    function addMainCss() {
        clearCss();

        const css = `
            body {
                font-size: 13px !important;
            }

            button {
                font-family: Arial !important;
                font-size: 13px !important;
            }

            select {
                font-family: Arial !important;
                font-size: 13px !important;
            }

            textarea {
                font-family: Arial !important;
                font-size: 13px !important;
            }

            input {
                font-family: Arial !important;
                font-size: 13px !important;
            }

            .multiSelectInput {
                width: auto !important;
                margin: 4px 8px 4px 4px !important;
            }

            .multiSelectSubDiv {
                background: white;
                margin-left: 8px;
                padding: 4px 12px 4px 4px;
                display: flex;
                align-items: center;
                border: 1px;
                border-style: solid;
                color: black;
            }

            .multiSelectDiv {
                z-index: 40;
                position: relative !important;
                margin: 4px;
            }

            #dataTypeFilterFlyOut {
                border-color: black;
                background-color: white;
            }
            
            @media (prefers-color-scheme: dark) {
                #dataTypeFilterFlyOut {
                    border-color: #d3d3d3f2;
                    background-color: #28282B;
                }
            }

            #transParentOverlay {
                background-color:transparent;
                position:fixed;
                top:0;
                left:0;
                right:0;
                bottom:0;
                margin:0px;
                display:block;
                z-index: 20;
            }

            #relationshipBrowser {
                background-color:white;
                position:absolute;
                top:0;
                left:0;
                right:0;
                bottom:0;
                padding: 20px;
                margin:0px;
                display:none;
                z-index: 100;  
                height: fit-content;
            }
            
            @media (prefers-color-scheme: dark) {
                #relationshipBrowser {
                    background-color: #28282B;
                }
            }

            pre
            .string { color: firebrick; }
            .number { color: darkgreen; }
            .boolean { color: blue; }
            .null { color: magenta; }
            .guid { color: firebrick; }
            .pf-link { color: blue; }
            .primarykey { color: tomato; }

            @media (prefers-color-scheme: dark) {
            *:not(svg, .copyButton, path, a) {
                color: #d3d3d3f2;
            }

            .multiSelectSubDiv {
                background: #131313;
                color: #d3d3d3f2;
            }

            pre
            .string { color: #5cc3ed; }
            .number { color: #5bd75b; }
            .boolean { color: #5bd75b; }
            .null { color: #ae82eb; }
            .guid { color: #5cc3ed; }
            .pf-link { color: lightblue; }
            .primarykey { color: tomato; }
            
            body { 
                background: #28282B;
            }

            a {
                color: #E9E9E9;
            }

            button {
                background: #131313;
            }
            }

            .panel input {
                width: 300px;
                margin: 0 0 0 8px;
            }

            @media (prefers-color-scheme: dark) {
                .panel input {
                    background: #131313;
                    border-color: #18181a;
                }    
            }


            .panel textarea {
                width: 400px;
                margin: 0 0 0 20px;
            }

            @media (prefers-color-scheme: dark) {
                .panel textarea {
                    background: #131313;
                }    
            }

            .panel select {
                margin: 0 0 0 8px;
            }

            @media (prefers-color-scheme: dark) {
                .panel select {
                    background: #131313;
                }  
            }

            .panel span:not(.keySpan):not(.lookupField):not(.lookupEdit):not(.pf-link):not(.copiedNotification):not(.copyButton) {
                margin-right: 24px;
                padding-right: 16px;
            }

            .panel option:empty {
                display:none;
            }

            .panel .copyButton {
                color:dimgray;
                display: none;
                cursor: pointer;
            }  
            
            .field  {
                cursor: pointer;
            }    
            
            .panel .field:hover .copyButton {
                display: unset;
            }      


            @media (prefers-color-scheme: dark) {
                .panel .copyButton {
                    color:darkgray;
                }       
            }

            .panel .link {
                margin:0;
                padding:0;
            }

            .panel table {
                color: black;
                margin-left: 26px;
                border-collapse: collapse;
                border: 1px solid black;
                table-layout: fixed;
                width: 98%;
            }

            .panel thead td {
                font-weight: bold;
            }

            .panel td {
                padding: 4px;
                border: 1px solid;
                overflow:auto;
            }

            .panel .impersonationIdFieldLabel {
                padding:0px;
                margin:0px;
            }

            .panel {
                padding:10px;
                margin:0px;
                text-wrap:nowrap;
                display:block;
            }

            .bodyPreviewed {
                display: inline-flex;
                margin-top: 0px;
                margin-bottom: 0px;
            }

            .prePreviewed {
                width: 48vw;
                overflow-x: scroll !important;
                overflow-y: scroll !important;
                height: 100%;
                margin: 0px;
            }
                
            .monacoContainer {
                height: 94vh;
                width: 98vw;
                box-sizing: border-box;
            }

            .monacoActions {
                height: 2em;
                display: flex;
                align-items: center;
                border-top: 1px solid #aaa;
                padding: 0.2em;
                box-sizing: border-box;
            }
            
            .monacoLabel {
                padding-right: 0.3em;
            }

            .checkBoxDiv {
                display: flex;
                padding: 1px 0;
                align-items: center;
            }
            `

        addcss(css);
    }

    await makeItPretty();
})()