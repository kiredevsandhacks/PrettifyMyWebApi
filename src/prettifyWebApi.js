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
    const retrievedPluralNames = {};

    async function odataFetch(url) {
        const response = await fetch(url, { headers: { 'Prefer': 'odata.include-annotations="*"', 'Cache-Control': 'no-cache' } });

        return await response.json();
    }

    async function retrievePluralName(logicalName) {
        if (retrievedPluralNames.hasOwnProperty(logicalName)) {
            return retrievedPluralNames[logicalName];
        }

        const requestUrl = apiUrl + "EntityDefinitions?$select=EntitySetName&$filter=(LogicalName eq '" + logicalName + "')";

        const json = await odataFetch(requestUrl);

        const pluralName = json.value[0].EntitySetName;
        retrievedPluralNames[logicalName] = pluralName;

        return pluralName;
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

    async function retrieveUpdateableAttributes(logicalName) {
        const requestUrl = apiUrl + "EntityDefinitions(LogicalName='" + logicalName + "')/Attributes?$filter=IsValidForUpdate eq true";

        const json = await odataFetch(requestUrl);

        return json.value;
    }

    async function retrieveOptionSetMetadata(logicalName, fieldName) {
        const requestUrl = apiUrl + "EntityDefinitions(LogicalName='" + logicalName + "')/Attributes/Microsoft.Dynamics.CRM.PicklistAttributeMetadata?$select=LogicalName&$expand=OptionSet,GlobalOptionSet";

        const json = await odataFetch(requestUrl);

        return json.value;
    }

    async function retrieveBooleanFieldMetadata(logicalName, fieldName) {
        const requestUrl = apiUrl + "EntityDefinitions(LogicalName='" + logicalName + "')/Attributes/Microsoft.Dynamics.CRM.BooleanAttributeMetadata?$select=LogicalName&$expand=OptionSet,GlobalOptionSet";

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

        return `<a class='previewLink' data-pluralName='${escapeHtml(pluralName)}' data-guid='${escapeHtml(formattedGuid)}' href='#'>Preview</a>`;
    }

    async function generateEditAnchor(logicalName, guid) {
        const pluralName = await retrievePluralName(logicalName);
        const formattedGuid = guid.replace('{', '').replace('}', '');

        return `
<a class='editLink' data-logicalName='${escapeHtml(logicalName)}' data-pluralName='${escapeHtml(pluralName)}' data-guid='${escapeHtml(formattedGuid)}' href='#'>Edit this record</a>     
<div class='editMenuDiv' style='display: none;'>
    <div>    Bypass Custom Plugin Execution<input class='bypassPluginExecutionBox' type='checkbox' style='width:25px;'>
    </div><div>    Preview changes before committing save<input class='previewChangesBeforeSavingBox' type='checkbox' style='width:25px;' checked='true'>
    </div><div>    Impersonate another user<input class='impersonateAnotherUserCheckbox' type='checkbox' style='width:25px;'>
    </div><div class='impersonateDiv' style='display:none;'><div>      Base impersonation on this field: <select  class='impersonateAnotherUserSelect'><option value='systemuserid'>systemuserid</option><option value='azureactivedirectoryobjectid'>azureactivedirectoryobjectid</option></select>  <i><a href='https://learn.microsoft.com/en-us/power-apps/developer/data-platform/webapi/impersonate-another-user-web-api#how-to-impersonate-a-user' target='_blank'>What's this?</a></i>
    </div><div>      <span class='impersonationIdFieldLabel'>systemuserid:</span><input class='impersonateAnotherUserInput'>  <span class='impersonateUserPreview'></span>
    </div></div><div><div id='previewChangesDiv'></div>    <a class='submitLink' style='display: none;' href='#'>Save</a>
    </div>
</div>`.replaceAll('\n', '');
    }

    function createSpan(cls, value) {
        return `<span class='${escapeHtml(cls)} hover'>${escapeHtml(value)}<span class='copyButton'>` + clipBoardIcon + `</span></span>`;
    }

    function createLinkSpan(cls, value) {
        return `<span class='${escapeHtml(cls)}'>${value}</span>`;
    }

    function createFieldSpan(cls, value, fieldName) {
        return `<span style='display: inline-flex;' class='${escapeHtml(cls)} hover'>${escapeHtml(value)}<div class='inputContainer containerNotEnabled' style='display: none;' data-fieldName='${escapeHtml(fieldName)}'></div><span class='copyButton'>` + clipBoardIcon + `</span></span>`;
    }

    function createOptionSetSpan(cls, value, fieldName, formattedValue) {
        let insertedValue = '';

        // toString the value because it can be a number. The formattedValue is always a string
        if (value?.toString() !== formattedValue) {
            insertedValue = value + ' : ' + formattedValue;
        } else {
            insertedValue = value;
        }

        return `<span style='display: inline-flex;' class='${escapeHtml(cls)} hover'>${escapeHtml(insertedValue)}<div class='inputContainer containerNotEnabled' style='display: none;' data-fieldName='${escapeHtml(fieldName)}'></div><span class='copyButton'>` + clipBoardIcon + `</span></span>`;
    }

    async function enrichObjectWithHtml(jsonObj, logicalName, pluralName, primaryIdAttribute, isSingleRecord, isNested) {
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

                        ordered[key][nestedKey] = await enrichObjectWithHtml(nestedValue, null, null, null, null, true);
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
                ordered[key] = await enrichObjectWithHtml(value, null, null, null, null, true);
                continue;
            }

            if (isAnnotation(key)) {
                continue;
            }

            if (value != null && value.replaceAll) {
                value = value.replaceAll('"', replacedQuote);
                value = value.replaceAll(',', replacedComma);
            }

            if (keyHasLookupAnnotation(key, ordered)) {
                const formattedValueValue = ordered[key + formattedValueType];
                const navigationPropertyValue = ordered[key + navigationPropertyType];
                const lookupTypeValue = ordered[key + lookupType];

                const newApiUrl = await generateApiAnchorAsync(lookupTypeValue, value);
                const formUrl = generateFormUrlAnchor(lookupTypeValue, value);
                const previewUrl = await generatePreviewUrlAnchor(lookupTypeValue, value);

                ordered[key] = [
                    createLinkSpan('link', newApiUrl) + ' : ' +
                    createLinkSpan('link', formUrl) + ' : ' +
                    createLinkSpan('link', previewUrl),
                    createSpan(determineType(formattedValueValue), 'Name: ' + formattedValueValue),
                    createSpan(determineType(lookupTypeValue), 'LogicalName: ' + lookupTypeValue),
                    createSpan(determineType(navigationPropertyValue), 'NavigationProperty: ' + navigationPropertyValue)
                ];

                delete ordered[key + formattedValueType];
                delete ordered[key + navigationPropertyType];
                delete ordered[key + lookupType];
            } else if (keyHasFormattedValueAnnotation(key, ordered)) {
                ordered[key] = createOptionSetSpan(cls, value, key, ordered[key + formattedValueType]);
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

    function setInputMetadata(input, container, datatype) {
        input.classList.add('enabledInputField');
        input.dataset['fieldname'] = container.dataset.fieldname;
        input.dataset['datatype'] = datatype;

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

        for (let attribute of attributesMetadata) {
            for (let container of inputContainers) {
                if (container.dataset.fieldname === attribute.LogicalName) {
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
                    } else if (attributeType === 'Integer' || attributeType === 'Double') { // for now treat double the same as ints and let the server handle the validation
                        createInput(container, false, 'int');
                    } else if (attributeType === 'Decimal') {
                        createInput(container, false, 'decimal');
                    } else if (attributeType === 'Boolean') {
                        const fieldOptionSetMetadata = booleanMetadata.find(osv => osv.LogicalName === attribute.LogicalName);
                        if (fieldOptionSetMetadata) {
                            const fieldOptionset = fieldOptionSetMetadata.GlobalOptionSet || fieldOptionSetMetadata.OptionSet;
                            createBooleanInput(container, fieldOptionset.FalseOption, fieldOptionset.TrueOption)
                        }
                    } else if (attributeType === 'DateTime') {
                        // todo
                    } else if (attributeType === 'Uniqueidentifier') {
                        // can't change this
                    } else if (attributeType === 'State') {
                        // difficult to implement
                    } else if (attributeType === 'Status') {
                        // difficult to implement
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

        const submitLink = document.getElementsByClassName('mainPanel')[0].getElementsByClassName('submitLink')[0];
        submitLink.style.display = null;
        submitLink.onclick = async function () {
            await submitEdit(pluralName, id);
        }

        // remove all hover handlers as they mess up the foratting and are not wanted in the editing context
        Array.from(document.querySelectorAll('.hover')).forEach((el) => el.classList.remove('hover'));
    }

    async function submitEdit(pluralName, id) {
        const changedFields = {};
        const enabledFields = document.getElementsByClassName('mainPanel')[0].getElementsByClassName('enabledInputField');
        for (let input of enabledFields) {
            const originalValue = window.originalResponseCopy[input.dataset.fieldname];
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
            } else if (dataType === 'decimal') {
                if (inputValue == null || inputValue === undefined || inputValue === '') {
                    value = null;
                } else {
                    if (inputValue.includes(',')) {
                        alert(fieldName + ' is a decimal number and contains a comma (,). Use a dot (.) as the separator.');
                        return;
                    }

                    value = parseFloat(inputValue);
                    if (isNaN(value)) {
                        alert(fieldName + ' is a decimal number. The value ' + inputValue + ' is not compatible.');
                        return;
                    }

                    if (/^-?[0-9]\d*(\.\d+)?$/.test(inputValue) === false) {
                        alert(fieldName + ' is a decimal number. The value ' + inputValue + ' is not compatible.');
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

            if (value !== originalValue && !(value === '' && originalValue == null)) {
                if (dataType === 'memo') {
                    if (originalValue?.replaceAll('\r\n', '\n') !== value) {
                        changedFields[fieldName] = value;
                    }
                } else {
                    changedFields[fieldName] = value;
                }
            }
        }

        const previewChangesBeforeSaving = document.getElementsByClassName('mainPanel')[0].getElementsByClassName('previewChangesBeforeSavingBox')[0].checked;
        const submitLink = document.getElementsByClassName('mainPanel')[0].getElementsByClassName('submitLink')[0];

        if (!!previewChangesBeforeSaving) {
            previewChanges(changedFields, pluralName, id);
            submitLink.style.display = 'none';
            return;
        }

        await commitSave(pluralName, id, changedFields);
    }

    async function commitSave(pluralName, id, changedFields) {
        const requestUrl = apiUrl + pluralName + '(' + id + ')';

        const headers = {
            'accept': 'application/json',
            'content-type': 'application/json',
            'OData-MaxVersion': '4.0',
            'OData-Version': '4.0',
            'If-Match': '*'
        };

        const bypassCustomPluginExecution = document.getElementsByClassName('bypassPluginExecutionBox')[0].checked;
        if (!!bypassCustomPluginExecution) {
            headers['MSCRM.BypassCustomPluginExecution'] = true;
        }

        const impersonateAnotherUser = document.getElementsByClassName('impersonateAnotherUserCheckbox')[0].checked;
        const impersonateAnotherUserField = document.getElementsByClassName('impersonateAnotherUserSelect')[0].value;
        const impersonateAnotherUserInput = document.getElementsByClassName('impersonateAnotherUserInput')[0].value;

        if (!!impersonateAnotherUser) {
            if (impersonateAnotherUserInput == null || impersonateAnotherUserInput == '') {
                alert('User impersonation was checked, but ' + impersonateAnotherUserField + ' is empty');
                return;
            }

            if (impersonateAnotherUserInput?.length != 36) {
                alert('Error while impersonating user: ' + impersonateAnotherUserInput + ' is not a valid guid.');
                return;
            }

            if (impersonateAnotherUserField == 'systemuserid') {
                headers['MSCRMCallerID'] = impersonateAnotherUserInput;
            } else if (impersonateAnotherUserField == 'azureactivedirectoryobjectid') {
                headers['CallerObjectId'] = impersonateAnotherUserInput;
            } else {
                alert('This should not happen. Wrong value in impersonateAnotherUserSelect: ' + impersonateAnotherUserField);
                return;
            }
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
                jsonObj[valueKeyWithCount][key] = await enrichObjectWithHtml(jsonObj[valueKeyWithCount][key], result.logicalName, pluralName, result.primaryIdAttribute, false, false);
            }
        } else {
            if (generateEditLink) {
                window.originalResponseCopy = JSON.parse(JSON.stringify(jsonObj));
            }
            jsonObj = await enrichObjectWithHtml(jsonObj, result.logicalName, pluralName, result.primaryIdAttribute, generateEditLink, false);
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

        if (!isMultiple && generateEditLink) {
            setImpersonateUserHandlers();
        }
    }

    function previewChanges(changedFields, pluralName, id) {
        const changes = [];

        for (let key in changedFields) {
            const change = {};
            const originalValue = window.originalResponseCopy[key];
            const updatedValue = changedFields[key];

            change.column = key;
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
        undoAllLink.href = '#';

        undoAllLink.onclick = makeItPretty;

        editMenu.appendChild(undoAllLink);

        const lineBreak2 = document.createElement('br');
        editMenu.appendChild(lineBreak2);
        editMenu.append('    ');

        const submitChangesLink = document.createElement('a');
        submitChangesLink.innerText = 'Commit Save';
        submitChangesLink.href = '#';

        // create this callback so we enclose the values we need when saving
        const saveCallback = async function () {
            await commitSave(pluralName, id, changedFields);
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

            input, textarea {
              width: 300px;
              margin: 0 0 0 8px;
            }

            select {
                margin: 0 0 0 8px;
            }

            span {
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
            
            .hover:hover .copyButton {
              display: unset;
            }

            .copyButton:hover {
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