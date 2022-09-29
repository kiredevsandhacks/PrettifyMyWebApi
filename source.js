javascript: (async function () {
    const formattedValueType = '@OData.Community.Display.V1.FormattedValue';
    const navigationPropertyType = '@Microsoft.Dynamics.CRM.associatednavigationproperty';
    const lookupType = '@Microsoft.Dynamics.CRM.lookuplogicalname';

    const replacedQuote = '__~~__REPLACEDQUOTE__~~__';
    const replacedComma = '__~~__REPLACEDCOMMA__~~__';

    const apiUrl = /\/api\/data\/v[0-9][0-9]?.[0-9]\//.exec(window.location.pathname)[0];
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

    var entityMap = {
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

    function generateWebApiAnchor(guid) {
        const formattedGuid = guid.replace('{', '').replace('}', '');
        const newLocation = apiUrl + escapeHtml(window.currentEntityPluralName) + '(' + escapeHtml(formattedGuid) + ')';

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

        return `<a class='editLink' data-logicalName='${escapeHtml(logicalName)}' data-pluralName='${escapeHtml(pluralName)}' data-guid='${escapeHtml(formattedGuid)}' href='#'>Edit this record</a>     <div class='editMenuDiv' style='display: none;'><div>    Bypass Custom Plugin Execution<input class='bypassPluginExecutionBox' type='checkbox' style='width:25px;'></div><div>    Preview changes before committing save<input class='previewChangesBeforeSavingBox' type='checkbox' style='width:25px;' checked='true'></div><div><div id='previewChangesDiv'></div>    <a class='submitLink' style='display: none;' href='#'>Save</a></div></div>  `;
    }

    function createSpan(cls, value) {
        return `<span class='${escapeHtml(cls)}'>${escapeHtml(value)}</span>`;
    }

    function createLinkSpan(cls, value) {
        return `<span class='${escapeHtml(cls)}'>${value}</span>`;
    }

    function createFieldSpan(cls, value, fieldName) {
        return `<span style='display: inline-flex;' class='${escapeHtml(cls)}'>${escapeHtml(value)}<div class='inputContainer containerNotEnabled' style='display: none;' data-fieldName='${escapeHtml(fieldName)}'></div></span>`;
    }

    function createOptionSetSpan(cls, value, fieldName, formattedValue) {
        const insertedValue = value + ' : ' + formattedValue;
        return `<span style='display: inline-flex;' class='${escapeHtml(cls)}'>${escapeHtml(insertedValue)}<div class='inputContainer containerNotEnabled' style='display: none;' data-fieldName='${escapeHtml(fieldName)}'></div></span>`;
    }

    async function enrichObjectWithHtml(jsonObj, logicalName, primaryIdAttribute, isSingleRecord) {
        const recordId = jsonObj[primaryIdAttribute]; // we need to get this value before parsing or else it will contain html

        const ordered = Object.keys(jsonObj).sort(
            (obj1, obj2) => {
                let obj1Underscore = obj1.startsWith('_');
                let obj2Underscore = obj2.startsWith('_');
                if (obj1Underscore && !obj2Underscore) {
                    return 1;
                } else if (!obj1Underscore && obj2Underscore) {
                    return -1;
                }

                else return obj1 > obj2 ? 1 : -1;

            }).reduce(
                (obj, key) => {
                    obj[key] = jsonObj[key];
                    return obj;
                },
                {}
            );

        for (let key in ordered) {
            let value = ordered[key];

            if (typeof (value) === "object" && value != null) {
                await enrichObjectWithHtml(value);
                continue;
            }

            if (isAnnotation(key)) {
                continue;
            }

            if (value != null && value.replaceAll) {
                value = value.replaceAll('"', replacedQuote);
                value = value.replaceAll(',', replacedComma);
            }

            const cls = determineType(value);

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
                    createSpan(determineType(formattedValueValue), formattedValueValue),
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
                    ordered[key] = createSpan('primarykey', value);
                } else {
                    ordered[key] = createFieldSpan(cls, value, key);
                }
            }
        }

        const newObj = {};
        if (logicalName != null && logicalName !== '' && recordId != null && recordId !== '') {
            newObj["Form Link"] = createLinkSpan('link', generateFormUrlAnchor(logicalName, recordId));

            if (isSingleRecord) {
                newObj["Edit this record"] = createLinkSpan('link', await generateEditAnchor(logicalName, recordId));
            } else {
                newObj["Web Api Link"] = createLinkSpan('link', generateWebApiAnchor(recordId));
            }
        } else {
            newObj["Form Link"] = "Could not generate link";
            newObj["Web Api Link"] = "Could not generate link";
        }

        const combinedJsonObj = Object.assign(newObj, ordered);
        return combinedJsonObj;
    }

    function setPreviewLinkClickHandlers() {
        const previewLinks = document.getElementsByClassName('previewLink');

        for (let previewLink of previewLinks) {
            const pluralName = previewLink.attributes["data-pluralName"].value;
            const newLocation = pluralName + "(" + previewLink.attributes["data-guid"].value + ")";

            previewLink.onclick = function () {
                previewRecord(pluralName, newLocation);
            }
        }
    }

    function setEditLinkClickHandlers() {
        const editLinks = document.getElementsByClassName('editLink');

        for (let editLink of editLinks) {
            const logicalName = editLink.attributes["data-logicalName"].value;
            const pluralName = editLink.attributes["data-pluralName"].value;
            const id = editLink.attributes["data-guid"].value;

            editLink.onclick = async function () {
                await editRecord(logicalName, pluralName, id);
            }
        }
    }

    function createInput(container, multiLine, datatype) {
        const value = window.originalResponseCopy[container.dataset.fieldname];

        let input;

        if (!multiLine) {
            input = document.createElement("input");
        } else {
            input = document.createElement("textarea");
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
        input.classList.add("enabledInputField");
        input.dataset["fieldname"] = container.dataset.fieldname;
        input.dataset["datatype"] = datatype;

        container.parentElement.append(input);

        if (datatype === 'memo') {
            container.parentElement.style.display = null;
        }

        container.style.display = null;
        container.classList.remove("containerNotEnabled");
        container.classList.add("containerEnabled");
    }

    async function editRecord(logicalName, pluralName, id) {
        const editLink = document.getElementsByClassName('editLink')[0];
        editLink.style.display = 'none';

        const attributesMetadata = await retrieveUpdateableAttributes(logicalName);

        const optionSetMetadata = await retrieveOptionSetMetadata(logicalName);
        const booleanMetadata = await retrieveBooleanFieldMetadata(logicalName);

        const inputContainers = document.getElementsByClassName("mainPanel")[0].getElementsByClassName('inputContainer');

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

        const previewChangesBeforeSaving = document.getElementsByClassName('mainPanel')[0].getElementsByClassName('previewChangesBeforeSavingBox')[0].checked;

        const submitLink = document.getElementsByClassName('mainPanel')[0].getElementsByClassName('submitLink')[0];
        submitLink.style.display = null;
        submitLink.onclick = async function () {
            await submitEdit(pluralName, id);

            if (!!previewChangesBeforeSaving) {
                submitLink.style.display = 'none';
            }
        }
    }

    async function submitEdit(pluralName, id) {
        const changedFields = {};
        const enabledFields = document.getElementsByClassName('mainPanel')[0].getElementsByClassName('enabledInputField');
        for (let input of enabledFields) {
            const originalValue = window.originalResponseCopy[input.dataset.fieldname];
            const dataType = input.dataset.datatype;
            const inputValue = input.value;

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
                    alert('there was an error parsing the field ' + input.dataset.fieldname);
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
                        alert(input.dataset.fieldname + ' is a whole number. The value ' + inputValue + ' is not compatible.');
                        return;
                    }

                    if (/^-?\d+$/.test(inputValue) === false) {
                        alert(input.dataset.fieldname + ' is a whole number. The value ' + inputValue + ' is not compatible.');
                        return;
                    }
                }
            } else if (dataType === 'decimal') {
                if (inputValue == null || inputValue === undefined || inputValue === '') {
                    value = null;
                } else {
                    if (inputValue.includes(',')) {
                        alert(input.dataset.fieldname + ' is a decimal number and contains a comma (,). Use a dot (.) as the separator.');
                        return;
                    }

                    value = parseFloat(inputValue);
                    if (isNaN(value)) {
                        alert(input.dataset.fieldname + ' is a decimal number. The value ' + inputValue + ' is not compatible.');
                        return;
                    }

                    if (/^-?[0-9]\d*(\.\d+)?$/.test(inputValue) === false) {
                        alert(input.dataset.fieldname + ' is a decimal number. The value ' + inputValue + ' is not compatible.');
                        return;
                    }
                }
            } else if (dataType === 'bool') {
                if (!inputValue) {
                    // the select needs to contain a value always, if not, an error happened
                    alert('there was an error parsing the field ' + input.dataset.fieldname);
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
                        alert('there was an error parsing the field ' + input.dataset.fieldname);
                        return;
                    }
                }
            }

            if (value !== originalValue && !(value === '' && originalValue == null)) {
                changedFields[input.dataset.fieldname] = value;
            }
        }

        const previewChangesBeforeSaving = document.getElementsByClassName('mainPanel')[0].getElementsByClassName('previewChangesBeforeSavingBox')[0].checked;

        if (!!previewChangesBeforeSaving) {
            previewChanges(changedFields, pluralName, id);
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

        const response = await fetch(requestUrl, {
            method: 'PATCH',
            headers: headers,
            body: JSON.stringify(changedFields)
        });

        if (response.ok) {
            await start();
        } else {
            const errorText = await response.text();
            console.error(errorText);
            window.alert(errorText);
        }
    }

    function stringContains(str, value) {
        return str.indexOf(value) !== -1;
    }

    function isAnnotation(key) {
        return stringContains(key, formattedValueType) || stringContains(key, navigationPropertyType) || stringContains(key, lookupType);
    }

    async function replaceJsonAsync(jsonObj, htmlElement, pluralName, generateEditLink) {
        const isMultiple = (jsonObj.value && Array.isArray(jsonObj.value));

        const result = await retrieveLogicalNameFromPluralNameAsync(pluralName);

        if (isMultiple) {
            const valueKeyWithCount = 'value (' + jsonObj.value.length + ' records)';

            jsonObj[valueKeyWithCount] = jsonObj.value;
            delete jsonObj.value;

            for (const key in jsonObj[valueKeyWithCount]) {
                jsonObj[valueKeyWithCount][key] = await enrichObjectWithHtml(jsonObj[valueKeyWithCount][key], result.logicalName, result.primaryIdAttribute, false);
            }
        } else {
            if (generateEditLink) {
                window.originalResponseCopy = JSON.parse(JSON.stringify(jsonObj));
            }
            jsonObj = await enrichObjectWithHtml(jsonObj, result.logicalName, result.primaryIdAttribute, generateEditLink);
        }

        let json = JSON.stringify(jsonObj, undefined, 2);

        json = json.replaceAll('"', '').replaceAll(replacedQuote, escapeHtml('"'));
        json = json.replaceAll(',', '').replaceAll(replacedComma, ',');

        htmlElement.innerText = '';
        var pre = document.createElement('pre');
        if (generateEditLink) {
            pre.classList.add('mainPanel');
        }
        htmlElement.appendChild(pre).innerHTML = json;
        setPreviewLinkClickHandlers();
        setEditLinkClickHandlers();
    }

    function previewChanges(changedFields, pluralName, id) {
        const previewContainer = document.createElement('textarea');
        previewContainer.style.width = '500px';
        previewContainer.style.height = '300px';
        previewContainer.disabled = true;

        let previewText = '';
        for (let key in changedFields) {
            const originalValue = window.originalResponseCopy[key];
            const updatedValue = changedFields[key];
            previewText += key + ': \n';
            previewText += 'old: "' + originalValue + '"\n';
            if (updatedValue === null) { // literal null check. Remove double quotes for clear communication to the user that is an actuall null and not the string "null"
                previewText += 'new: ' + updatedValue + '\n';
            } else {
                previewText += 'new: "' + updatedValue + '"\n';
            }
        }

        // disable all stuff to prevent edits after previewing
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

        previewContainer.value = previewText;

        const editMenu = document.getElementById('previewChangesDiv');
        editMenu.innerHTML = '  ';
        editMenu.appendChild(previewContainer);

        const lineBreak = document.createElement('br');
        editMenu.appendChild(lineBreak);
        editMenu.append('    ');

        const undoAllLink = document.createElement('a');
        undoAllLink.innerText = 'Cancel';
        undoAllLink.href = '#';

        undoAllLink.onclick = start;

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

        await replaceJsonAsync(response, newDiv, pluralName, false);

        const btn = document.createElement('button');
        btn.style = `
            height: 30px;
            width: 30px;
            margin: 24px;
            position: absolute;
            right: 30px;
            background: transparent;
            border: transparent;
            cursor: pointer;
            padding:0;
            font-size:24;
            `

        btn.innerHTML = '<span>×</span>';

        btn.addEventListener('click', function () {
            if (document.getElementsByClassName('previewPanel').length === 1) {
                resetCSS();
            }

            newDiv.remove();
        });

        newDiv.firstChild.insertBefore(btn, newDiv.firstChild.firstChild);

        document.body.scrollLeft = Number.MAX_SAFE_INTEGER;
    }

    async function start() {
        const response = await odataFetch(window.location.href);

        window.currentEntityPluralName = window.location.pathname.split('/').pop().split('(').shift();

        resetCSS();

        await replaceJsonAsync(response, document.body, window.currentEntityPluralName, true);
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

            option:empty {
              display:none;
            }
            `

        addcss(css);
    }

    await start();
}())