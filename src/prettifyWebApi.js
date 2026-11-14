(async function () {
    if (!['#p', '#pf', '#pr'].includes(window.location.hash)) {
        return;
    }

    // ---------------------------------------------------------------------
    // Constants
    // ---------------------------------------------------------------------

    const formattedValueType = '@OData.Community.Display.V1.FormattedValue';
    const navigationPropertyType = '@Microsoft.Dynamics.CRM.associatednavigationproperty';
    const lookupType = '@Microsoft.Dynamics.CRM.lookuplogicalname';
    const metaDataKey = '@Microsoft.Dynamics.CRM.globalmetadataversion';
    const contextKey = '@odata.context';
    const etagKey = '@odata.etag';
    const totalRecordCountKey = '@Microsoft.Dynamics.CRM.totalrecordcount';
    const totalRecordCountLimitExceededKey = '@Microsoft.Dynamics.CRM.totalrecordcountlimitexceeded';

    const annotationTypes = [formattedValueType, navigationPropertyType, lookupType];
    const hiddenKeys = [contextKey, metaDataKey, etagKey];
    const hiddenCollectionKeys = [...hiddenKeys, totalRecordCountKey, totalRecordCountLimitExceededKey];

    const emptyGuid = '00000000-0000-0000-0000-000000000000';
    const settingsStorageKey = 'pfwaSettings';
    const impersonationDocsUrl = 'https://learn.microsoft.com/en-us/power-apps/developer/data-platform/webapi/impersonate-another-user-web-api#how-to-impersonate-a-user';

    const odataJsonHeaders = {
        'accept': 'application/json',
        'content-type': 'application/json',
        'OData-MaxVersion': '4.0',
        'OData-Version': '4.0',
    };

    const svgNamespace = 'http://www.w3.org/2000/svg';
    const clipboardIconPath = 'M19 2h-4.18C14.4.84 13.3 0 12 0S9.6.84 9.18 2H5c-1.1 0-2 .9-2 2v16c0 1.1.9 2 2 2h14c1.1 0 2-.9 2-2V4c0-1.1-.9-2-2-2zm-7 0c.55 0 1 .45 1 1s-.45 1-1 1s-1-.45-1-1s.45-1 1-1zm7 18H5V4h2v3h10V4h2v16z';
    const clipboardCopiedIconPath = 'M19 3h-4.18C14.4 1.84 13.3 1 12 1c-1.3 0-2.4.84-2.82 2H5c-1.1 0-2 .9-2 2v14c0 1.1.9 2 2 2h14c1.1 0 2-.9 2-2V5c0-1.1-.9-2-2-2zm-7 0c.55 0 1 .45 1 1s-.45 1-1 1s-1-.45-1-1s.45-1 1-1zm-2 14l-4-4l1.41-1.41L10 14.17l6.59-6.59L18 9l-8 8z';

    // maps Dataverse attribute types to the data type of a plain text input
    const textInputTypes = {
        String: 'string',
        EntityName: 'string',
        Memo: 'memo',
        Integer: 'int',
        Decimal: 'decimal',
        Money: 'money',
        Double: 'float',
        DateTime: 'datetime',
        Uniqueidentifier: 'uid',
    };

    const relationshipFilterFields = {
        manyToMany: ['Entity1LogicalName', 'Entity2LogicalName', 'IntersectEntityName', 'Entity1IntersectAttribute', 'Entity2IntersectAttribute', 'Entity1NavigationPropertyName', 'Entity2NavigationPropertyName', 'SchemaName'],
        oneToMany: ['ReferencedAttribute', 'ReferencedEntity', 'ReferencingAttribute', 'ReferencingEntity', 'ReferencedEntityNavigationPropertyName', 'ReferencingEntityNavigationPropertyName', 'SchemaName'],
    };

    let apiUrl = '';

    try {
        apiUrl = /([\/a-zA-Z0-9]+)?\/api\/data\/v[0-9][0-9]?.[0-9]\//.exec(window.location.pathname)[0];
    } catch {
        alert('It seems you are not viewing a form or the dataverse odata web api. If you think this is an error, please contact the author of the extension and he will fix it asap.');
        return;
    }

    // ---------------------------------------------------------------------
    // State
    // ---------------------------------------------------------------------

    const state = {
        mode: 'read', // 'read' | 'update' | 'create'
        originalRecord: null, // untouched copy of the response of the main panel
        currentPluralName: null,
        originalPluralName: null, // the plural name of the parent row when browsing a many to one relationship
        isManyToOneContext: false,
        titleSet: false,
        copyBusy: false,
        foundNavigationProperties: [],
        mainView: null,
    };

    const memoizedCalls = {};

    // A user facing validation error. Shown as an alert, after which the edit controls are reset.
    class UserError extends Error { }

    // ---------------------------------------------------------------------
    // DOM helpers
    // ---------------------------------------------------------------------

    // Creates an element. Children can be nodes, strings (become text nodes), arrays, or null/false (skipped).
    // Options: className, dataset, style, attributes, on (event handlers), and any other element property.
    function el(tag, options = {}, ...children) {
        const { className, dataset, style, attributes, on, ...properties } = options;
        const element = document.createElement(tag);

        if (className) {
            element.className = className;
        }
        if (dataset) {
            Object.assign(element.dataset, dataset);
        }
        if (style) {
            Object.assign(element.style, style);
        }
        for (const [name, value] of Object.entries(attributes ?? {})) {
            element.setAttribute(name, value);
        }
        for (const [eventName, handler] of Object.entries(on ?? {})) {
            element.addEventListener(eventName, handler);
        }

        appendChildren(element, children);

        // assign properties after the children, so that select.value can match the appended options
        Object.assign(element, properties);

        return element;
    }

    function appendChildren(parent, children) {
        for (const child of children.flat(Infinity)) {
            if (child != null && child !== false) {
                parent.append(child);
            }
        }
        return parent;
    }

    function setVisible(element, visible) {
        element.classList.toggle('pfwa-hidden', !visible);
    }

    // An anchor that runs a handler instead of navigating
    function actionLink(text, onClick, className) {
        return el('a', {
            className,
            href: '#',
            on: {
                click: (e) => {
                    e.preventDefault();
                    onClick(e);
                }
            }
        }, text);
    }

    function linkSpan(...children) {
        return el('span', { className: 'pf-link' }, children);
    }

    function createSvgIcon(className, pathData) {
        const svg = document.createElementNS(svgNamespace, 'svg');
        svg.setAttribute('class', className);
        svg.setAttribute('viewBox', '0 0 24 24');

        const path = document.createElementNS(svgNamespace, 'path');
        path.setAttribute('fill', 'currentColor');
        path.setAttribute('d', pathData);

        svg.appendChild(path);
        return svg;
    }

    function createToolbarButton(label, onClick) {
        return el('button', { type: 'button', on: { click: onClick } }, label);
    }

    // a row of the json-like output, e.g. 'key: value'. Pass null as key for array items.
    function createRow(key, ...content) {
        const keyNodes = key == null ? null : [el('span', { className: 'keySpan' }, key), ': '];
        return el('div', { className: 'pfwa-row' }, keyNodes, content);
    }

    // an indented block of rows, e.g. the contents of an object
    function createBlock(rows) {
        return el('div', { className: 'pfwa-block' }, rows);
    }

    // created once and cloned for every field
    const copyButtonTemplate = el('span', { className: 'copyButton' }, createSvgIcon('copyIcon', clipboardIconPath));
    const copiedNotificationTemplate = el('span', { className: 'copiedNotification' },
        createSvgIcon('copiedIcon', clipboardCopiedIconPath),
        el('span', { className: 'copiedText' }, 'Copied!'));

    // ---------------------------------------------------------------------
    // Web Api
    // ---------------------------------------------------------------------

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

    async function alertErrorResponse(response) {
        const errorText = await response.text();
        console.error(`${response.status} - ${errorText}`);
        window.alert(`${response.status} - ${errorText}`);
    }

    // returns true when the request succeeded, alerts the error otherwise
    async function sendRequest(url, method, body) {
        const options = { method };

        if (body !== undefined) {
            options.headers = { ...odataJsonHeaders, 'If-Match': '*' };
            options.body = JSON.stringify(body);
        }

        const response = await fetch(url, options);

        if (!response.ok) {
            await alertErrorResponse(response);
            return false;
        }
        return true;
    }

    // ---------------------------------------------------------------------
    // Metadata
    // ---------------------------------------------------------------------

    async function retrieveEntityDefinitions() {
        const json = await odataFetch(apiUrl + 'EntityDefinitions?$select=EntitySetName,LogicalName,PrimaryIdAttribute,PrimaryNameAttribute', true);

        return json.value;
    }

    async function findEntityByLogicalName(logicalName) {
        const entities = await retrieveEntityDefinitions();

        return entities.find(e => e.LogicalName === logicalName);
    }

    async function findEntityByPluralName(pluralName) {
        const entities = await retrieveEntityDefinitions();
        const lowerCasePluralName = pluralName?.toLowerCase();

        return entities.find(e => e.EntitySetName === pluralName)
            ?? entities.find(e => e.EntitySetName?.toLowerCase() === lowerCasePluralName);
    }

    async function retrievePluralName(logicalName) {
        return (await findEntityByLogicalName(logicalName))?.EntitySetName;
    }

    async function retrieveLogicalName(pluralName) {
        return (await findEntityByPluralName(pluralName))?.LogicalName;
    }

    // returns an empty object when the plural name does not belong to a table, e.g. for sdk messages
    async function retrieveEntityInfo(pluralName) {
        const entity = await findEntityByPluralName(pluralName);

        if (entity == null) {
            return {};
        }

        return {
            logicalName: entity.LogicalName,
            primaryIdAttribute: entity.PrimaryIdAttribute,
            primaryNameAttribute: entity.PrimaryNameAttribute
        };
    }

    // validityProperty: IsValidForUpdate, IsValidForCreate
    async function retrieveAttributes(logicalName, validityProperty) {
        const json = await odataFetch(`${apiUrl}EntityDefinitions(LogicalName='${logicalName}')/Attributes?$filter=${validityProperty} eq true`, true);

        return json.value;
    }

    // metadataType: PicklistAttributeMetadata, MultiSelectPicklistAttributeMetadata, StateAttributeMetadata, etc.
    async function retrieveOptionSetAttributes(logicalName, metadataType) {
        const json = await odataFetch(`${apiUrl}EntityDefinitions(LogicalName='${logicalName}')/Attributes/Microsoft.Dynamics.CRM.${metadataType}?$select=LogicalName&$expand=OptionSet,GlobalOptionSet`, true);

        return json.value;
    }

    async function retrieveEditableAttributes(logicalName, isCreateMode) {
        const updateableAttributes = await retrieveAttributes(logicalName, 'IsValidForUpdate');

        if (!isCreateMode) {
            return updateableAttributes;
        }

        const creatableAttributes = [...await retrieveAttributes(logicalName, 'IsValidForCreate')];

        // the statecode is creatable but not listed as such, so add it 'manually'
        const stateCodeAttribute = updateableAttributes.find(a => a.LogicalName === 'statecode');
        if (stateCodeAttribute != null && !creatableAttributes.some(a => a.LogicalName === 'statecode')) {
            creatableAttributes.push(stateCodeAttribute);
        }

        return creatableAttributes;
    }

    async function retrieveEntityClientMetadata(logicalName) {
        const requestUrl = apiUrl + "GetClientMetadata(ClientMetadataQuery=@ClientMetadataQuery)?@ClientMetadataQuery={'MetadataType':'entity','MetadataSubtype':'{\"" + logicalName + "\":[\"merged\"]}'}";

        const json = await odataFetch(requestUrl, true);

        return JSON.parse(json.Metadata);
    }

    async function retrieveNavigationProperty(lookupEntitylogicalName, fieldName) {
        const currentRecordLogicalName = await retrieveLogicalName(state.currentPluralName);

        const metadata = await retrieveEntityClientMetadata(currentRecordLogicalName);

        if (metadata.Entities.length != 1) {
            alert(`Something went wrong with retrieving the navigation property for ${currentRecordLogicalName}/${lookupEntitylogicalName}/${fieldName}: ${metadata.Entities.length} entities found. This should not happen.`);
            return null;
        }

        const relations = metadata.Entities[0].ManyToOneRelationships.filter(rel => fieldName === 'ownerid'
            ? rel.ReferencingAttribute === fieldName
            : rel.ReferencingEntity === currentRecordLogicalName && rel.ReferencedEntity === lookupEntitylogicalName && rel.ReferencingAttribute === fieldName);

        if (relations.length != 1) {
            alert(`Something went wrong with retrieving the navigation property for ${currentRecordLogicalName}/${lookupEntitylogicalName}/${fieldName}: ${relations.length} relations found. This should not happen.`);
            return null;
        }

        return relations[0].ReferencingEntityNavigationPropertyName;
    }

    async function retrieveRelationships() {
        const currentRecordLogicalName = await retrieveLogicalName(state.currentPluralName);

        const metadata = await retrieveEntityClientMetadata(currentRecordLogicalName);

        if (metadata.Entities.length != 1) {
            return null;
        }

        const { ManyToManyRelationships, OneToManyRelationships, ManyToOneRelationships } = metadata.Entities[0];

        return { ManyToManyRelationships, OneToManyRelationships, ManyToOneRelationships };
    }

    async function retrieveRelationshipDefinition(relationShipName) {
        const relationships = await retrieveRelationships();
        const combined = [...relationships.ManyToManyRelationships, ...relationships.ManyToOneRelationships, ...relationships.OneToManyRelationships];
        const matches = combined.filter(r =>
            r.SchemaName === relationShipName ||
            r.Entity1NavigationPropertyName === relationShipName ||
            r.Entity2NavigationPropertyName === relationShipName ||
            r.ReferencedEntityNavigationPropertyName === relationShipName ||
            r.ReferencingEntityNavigationPropertyName === relationShipName
        );

        return matches.length === 1 ? matches[0] : null;
    }

    function findOptionSet(attributesMetadata, logicalName) {
        const attributeMetadata = attributesMetadata.find(a => a.LogicalName === logicalName);

        return attributeMetadata ? (attributeMetadata.GlobalOptionSet || attributeMetadata.OptionSet) : null;
    }

    // ---------------------------------------------------------------------
    // Value helpers
    // ---------------------------------------------------------------------

    function stripBraces(guid) {
        return guid?.replace('{', '')?.replace('}', '');
    }

    function isNumber(n) {
        return typeof n == 'number' && !isNaN(n) && isFinite(n);
    }

    function determineType(value) {
        if (isNumber(value)) {
            return 'number';
        } else if (value === null) { // null check is explicit here (===) to prevent matching undefined (which renders as string)
            return 'null';
        } else if (typeof value === 'boolean') {
            return 'boolean';
        }
        return 'string';
    }

    function isAnnotation(key) {
        return annotationTypes.some(type => key.includes(type));
    }

    function isObject(value) {
        return value !== null && typeof value === 'object';
    }

    // values are rendered on a single line, so show line breaks as \n
    function formatDisplayText(value) {
        return String(value).replaceAll('\r\n', '\\n').replaceAll('\n', '\\n');
    }

    // underscore prefixed keys (lookups) last, the rest alphabetically
    function sortKeys(keys) {
        return [...keys].sort((key1, key2) => {
            const key1Underscore = key1.startsWith('_');
            const key2Underscore = key2.startsWith('_');
            if (key1Underscore && !key2Underscore) {
                return 1;
            } else if (!key1Underscore && key2Underscore) {
                return -1;
            }
            return key1 > key2 ? 1 : -1;
        });
    }

    // _parentaccountid_value => parentaccountid
    function lookupFieldName(key) {
        return key.replace(/^_/, '').replace(/_value$/, '');
    }

    // ---------------------------------------------------------------------
    // Links
    // ---------------------------------------------------------------------

    function createApiLink(pluralName, guid, text) {
        return el('a', { target: '_blank', href: `${apiUrl}${pluralName}(${stripBraces(guid)})#p` }, text);
    }

    function createWebApiLink(guid, pluralName) {
        return createApiLink(pluralName, guid, 'Open in Web Api');
    }

    function createFormLink(logicalName, guid) {
        return el('a', { target: '_blank', href: `/main.aspx?etn=${logicalName}&id=${guid}&pagetype=entityrecord` }, 'Open in Form');
    }

    function createDisassociateLink(relationship, pluralName, recordId) {
        const id = stripBraces(recordId);
        let disassociate = null;

        if (state.isManyToOneContext) {
            disassociate = () => disassociateRowManyToOne(relationship.ReferencingEntityNavigationPropertyName);
        } else if (relationship.RelationshipType === 1) {
            disassociate = () => disassociateRowManyToMany(pluralName, id);
        } else if (relationship.RelationshipType === 0) {
            disassociate = () => disassociateRowOneToMany(pluralName, id, relationship.ReferencingEntityNavigationPropertyName);
        } else {
            return null;
        }

        const link = actionLink('Disassociate this row', async () => {
            setVisible(link, false);
            await disassociate();
            setVisible(link, true);
        });

        return link;
    }

    // ---------------------------------------------------------------------
    // Fields
    // ---------------------------------------------------------------------

    // a value which can be copied to the clipboard by clicking it
    function createField(cls, value, ...extraClasses) {
        const text = formatDisplayText(value);
        const field = el('span', { className: [cls, 'field', ...extraClasses].join(' ') },
            text,
            copyButtonTemplate.cloneNode(true),
            copiedNotificationTemplate.cloneNode(true));

        field.addEventListener('click', () => copyFieldToClipboard(field, text));

        return field;
    }

    // a field which gets an input appended when the row is edited
    function createEditableField(cls, value, fieldName, view) {
        const field = createField(cls, value);

        view?.slots.set(fieldName, { kind: 'field', fieldName, element: field });

        return field;
    }

    // the container which receives the lookup edit controls when the row is edited
    function createLookupSlot(view, key, guid, lookupLogicalName) {
        const element = el('span', { className: 'lookupEdit' });
        const fieldName = lookupFieldName(key);

        view?.slots.set(fieldName, {
            kind: 'lookup',
            fieldName,
            element,
            id: stripBraces(guid) ?? null,
            lookupType: lookupLogicalName ?? null
        });

        return element;
    }

    async function copyFieldToClipboard(field, text) {
        // no copying while editing, and prevent weird issues with double clicking
        if (field.closest('.editing') || state.copyBusy) {
            return;
        }

        state.copyBusy = true;

        try {
            await navigator.clipboard.writeText(text);

            field.classList.add('copied');

            setTimeout(() => {
                field.classList.remove('copied');
                state.copyBusy = false;
            }, 400);
        } catch {
            state.copyBusy = false;
            alert('Failed to copy');
        }
    }

    // ---------------------------------------------------------------------
    // Rendering
    // ---------------------------------------------------------------------

    /*
     * Render options:
     *   logicalName, pluralName, primaryIdAttribute, primaryNameAttribute: metadata of the rendered table (null for nested objects)
     *   isSingleRecord: render the edit and delete actions
     *   isNested: render no actions at all
     *   isCreateMode: render the 'create new row' action
     *   relationship: the relationship definition when browsing related rows
     *   isSingleColumnValueOnly: the response only contains a single column value
     *   view: the main view to register editable fields in, null when not editable
     */
    const nestedRenderOptions = { isNested: true };

    async function renderRecordRows(record, options) {
        const rows = options.isNested ? [] : await renderRecordActionRows(record, options);

        rows.push(...await renderPropertyRows(record, options));

        return rows;
    }

    async function renderRecordActionRows(record, options) {
        const { logicalName, pluralName, isSingleRecord, isCreateMode, relationship, isSingleColumnValueOnly, view } = options;
        const recordId = record[options.primaryIdAttribute];
        const hasRecordId = recordId != null && recordId !== '';
        const rows = [];

        if (logicalName && hasRecordId) {
            rows.push(createRow('Form Link', linkSpan(createFormLink(logicalName, recordId))));

            if (isSingleRecord) {
                const editPluralName = await retrievePluralName(logicalName);
                const id = stripBraces(recordId);

                rows.push(createRow('Edit this row', createEditMenu(view, { logicalName, pluralName: editPluralName, id, isCreateMode: false })));
                rows.push(createDeleteRow(editPluralName, id));
            } else {
                rows.push(createRow('Web Api Link', linkSpan(createWebApiLink(recordId, pluralName))));
            }

            const disassociateLink = relationship ? createDisassociateLink(relationship, pluralName, recordId) : null;
            if (disassociateLink) {
                rows.push(createRow('Disassociate this row', linkSpan(disassociateLink)));
            }
        } else if (!isSingleColumnValueOnly && !isCreateMode && logicalName && !hasRecordId) {
            rows.push(createRow('Form Link', 'Could not generate link'));
            rows.push(createRow('Web Api Link', 'Could not generate link'));
        }

        if (isCreateMode) {
            const createPluralName = await retrievePluralName(logicalName);
            rows.push(createRow('Create new row', createEditMenu(view, { logicalName, pluralName: createPluralName, id: null, isCreateMode: true })));
        }

        return rows;
    }

    function createDeleteRow(pluralName, id) {
        const deleteLink = actionLink('Delete this row', () => deleteRecord(pluralName, id, deleteLink));
        const row = createRow('Delete this row', linkSpan(deleteLink));
        row.classList.add('deleteRow');

        return row;
    }

    // nested objects and arrays are listed first, followed by the columns
    async function renderPropertyRows(record, options) {
        const nestedRows = [];
        const columnRows = [];
        const consumedKeys = new Set(); // annotations which are rendered as part of another key

        for (const key of sortKeys(Object.keys(record))) {
            if (consumedKeys.has(key) || hiddenKeys.includes(key)) {
                continue;
            }

            const value = record[key];

            if (isObject(value)) {
                nestedRows.push(createRow(key, await renderNestedValue(value)));
            } else if (isAnnotation(key)) {
                // an annotation without a matching column
                nestedRows.push(createRow(key, formatDisplayText(value)));
            } else {
                columnRows.push(createRow(key, await renderPropertyValue(record, key, options, consumedKeys)));
            }
        }

        return [...nestedRows, ...columnRows];
    }

    async function renderNestedValue(value) {
        if (!Array.isArray(value)) {
            return ['{', createBlock(await renderRecordRows(value, nestedRenderOptions)), '}'];
        }

        const items = [];
        for (const item of value) {
            items.push(createRow(null, isObject(item) ? await renderNestedValue(item) : createField(determineType(item), item)));
        }

        return ['[', createBlock(items), ']'];
    }

    async function renderPropertyValue(record, key, options, consumedKeys) {
        const value = record[key];
        const cls = determineType(value);
        const { view } = options;

        if (record[key + lookupType]) {
            [formattedValueType, navigationPropertyType, lookupType].forEach(type => consumedKeys.add(key + type));

            return await renderLookupValue(record, key, view);
        }

        if (key.startsWith('_') && key.endsWith('_value')) {
            // a lookup without a value
            return [createField(cls, value, 'lookupField'), createLookupSlot(view, key, null, null)];
        }

        const formattedValue = record[key + formattedValueType];
        if (formattedValue) {
            consumedKeys.add(key + formattedValueType);

            // toString the value because it can be a number. The formattedValue is always a string
            const text = value?.toString() !== formattedValue ? value + ' : ' + formattedValue : value;

            return createEditableField(cls, text, key, view);
        }

        if (options.logicalName !== 'solution' && key === 'solutionid' && value != null) {
            return [createEditableField(cls, value, key, view), createWebApiLink(value, 'solutions')];
        }

        if (key === options.primaryIdAttribute) {
            return el('b', {}, createField('primarykey', value));
        }

        if (key === options.primaryNameAttribute) {
            return el('b', {}, createEditableField(cls, value, key, view));
        }

        return createEditableField(cls, value, key, view);
    }

    async function renderLookupValue(record, key, view) {
        const guid = record[key];
        const formattedValue = record[key + formattedValueType];
        const navigationProperty = record[key + navigationPropertyType];
        const lookupLogicalName = record[key + lookupType];

        state.foundNavigationProperties.push(navigationProperty);

        const pluralName = await retrievePluralName(lookupLogicalName);
        const id = stripBraces(guid);

        const lookupDisplay = el('span', { className: 'lookupDisplay' },
            '{',
            createBlock([
                createRow(null,
                    linkSpan(createApiLink(pluralName, id, id)), ' - ',
                    linkSpan(createFormLink(lookupLogicalName, guid)), ' - ',
                    linkSpan(actionLink('Preview', () => previewRecord(pluralName, id)))),
                createRow(null, createField(determineType(formattedValue), 'Name: ' + formattedValue)),
                createRow(null, createField(determineType(lookupLogicalName), 'LogicalName: ' + lookupLogicalName)),
                createRow(null, createField(determineType(navigationProperty), 'NavigationProperty: ' + navigationProperty)),
            ]),
            '}');

        // in edit mode, only the name is displayed
        const lookupName = el('span', { className: 'lookupEdit' }, createField('string', formattedValue, 'lookupField'));

        return [lookupDisplay, lookupName, createLookupSlot(view, key, guid, lookupLogicalName)];
    }

    // Collections: the remaining top level keys (e.g. @odata.nextLink) followed by the rows
    async function renderCollectionRows(json, options) {
        const rows = [];

        for (const key of Object.keys(json)) {
            if (key === 'value' || hiddenCollectionKeys.includes(key)) {
                continue;
            }

            const value = json[key];
            rows.push(createRow(key, isObject(value) ? await renderNestedValue(value) : formatDisplayText(value)));
        }

        const records = [];
        for (const record of json.value) {
            records.push(createRow(null, '{', createBlock(await renderRecordRows(record, options)), '}'));
        }

        rows.push(createRow(`value (${json.value.length} rows)`, '[', createBlock(records), ']'));

        return rows;
    }

    // Determines whether the url browses a relationship, e.g. /accounts(id)/contact_customer_accounts
    // and if so, switches the context to the related table
    async function resolveRelationshipContext(pluralName, entity, isSingleColumnValueOnly) {
        const context = { relationship: null, pluralName, entity };

        // the second part of this statement is a bit weird
        // but we do this to check if the context is 'manually' expanded on a relationship
        const isInferred = !isSingleColumnValueOnly && location.pathname.includes(')/');
        if (location.hash !== '#pr' && !isInferred) {
            return context;
        }

        const relationshipName = location.pathname.split('/').pop();

        try {
            const relationship = await retrieveRelationshipDefinition(relationshipName);

            if (relationship == null) {
                throw `Relationship not found`;
            }

            const current = entity.logicalName;
            let relatedLogicalName = null;

            if (relationship.RelationshipType === 1) {
                const { Entity1LogicalName: first, Entity2LogicalName: second } = relationship;

                if (first !== current && second === current) {
                    relatedLogicalName = first;
                } else if (first === current && second !== current) {
                    relatedLogicalName = second;
                } else if (first !== current && second !== current) {
                    throw `Cannot map relationship to either ${first} or ${second}`;
                }
                // else: self referencing table. No need to override the context
            } else if (relationship.RelationshipType === 0) {
                const { ReferencingEntity: referencing, ReferencedEntity: referenced } = relationship;

                if (referencing !== current && referenced === current) {
                    relatedLogicalName = referencing;
                } else if (referencing === current && referenced !== current) {
                    relatedLogicalName = referenced;
                    state.originalPluralName = pluralName;
                    state.isManyToOneContext = true;
                } else if (referencing !== current && referenced !== current) {
                    throw `Cannot map relationship to either ${referencing} or ${referenced}`;
                }
                // else: self referencing table. No need to override the context
            } else {
                throw `Cannot handle RelationshipType ${relationship.RelationshipType}`;
            }

            context.relationship = relationship;

            const relatedPluralName = relatedLogicalName ? await retrievePluralName(relatedLogicalName) : null;
            if (relatedPluralName) {
                // override the context to be the related table
                context.pluralName = relatedPluralName;
                context.entity = await retrieveEntityInfo(relatedPluralName);
                state.currentPluralName = relatedPluralName;
            }
        } catch (err) {
            if (!isInferred) {
                alert(`Something went wrong with retrieving the relationship definition for '${relationshipName}'. Error message: ${err}`);
            }
        }

        return context;
    }

    function setTitle(title) {
        if (!state.titleSet) {
            document.title = title;
            state.titleSet = true;
        }
    }

    // Renders a Web Api response into the container. Returns the rendered panel and its toolbar.
    async function renderPanel(json, container, initialPluralName, { isPreview = false, isCreateMode = false } = {}) {
        if (!isPreview) {
            state.mode = isCreateMode ? 'create' : 'read';
            if (!isCreateMode) {
                state.originalRecord = structuredClone(json);
            }
        }

        const isMultiple = Array.isArray(json.value);
        const isSingleColumnValueOnly = Object.keys(json).length === 2 && Boolean(json[contextKey]) && 'value' in json;

        const initialEntity = await retrieveEntityInfo(initialPluralName);

        if (window.location.hash === '#pf' && json.value?.length === 1) {
            // a flow was opened from Power Automate, redirect to the single row
            const recordId = json.value[0][initialEntity.primaryIdAttribute];
            window.location.href = apiUrl + initialPluralName + '(' + recordId + ')#p';
            return null;
        }

        const { relationship, pluralName, entity } = await resolveRelationshipContext(initialPluralName, initialEntity, isSingleColumnValueOnly);
        const { logicalName } = entity;
        const view = isPreview ? null : { panel: null, slots: new Map(), editMenu: null, editors: [], lockedControls: [] };

        const recordOptions = {
            logicalName,
            pluralName,
            primaryIdAttribute: entity.primaryIdAttribute,
            primaryNameAttribute: entity.primaryNameAttribute,
            isSingleRecord: !isPreview,
            isNested: false,
            isCreateMode,
            relationship,
            isSingleColumnValueOnly,
            view: null,
        };

        let rows;
        let singleRecordId = null;

        if (!logicalName) {
            // sdk messages other than retrieve or retrievemultiple
            rows = await renderRecordRows(json, { ...recordOptions, relationship: null, isSingleColumnValueOnly: false });
        } else if (isMultiple) {
            setTitle(relationship != null ? 'related ' + pluralName : pluralName);
            rows = await renderCollectionRows(json, { ...recordOptions, isSingleRecord: false, isSingleColumnValueOnly: false });
        } else {
            setTitle(logicalName);
            singleRecordId = json[entity.primaryIdAttribute];
            rows = await renderRecordRows(json, { ...recordOptions, view });
        }

        const toolbar = el('div', { className: 'pfwa-toolbar' });
        const panel = el('pre', { className: isPreview ? 'pfwa-root' : 'pfwa-root mainPanel panel' }, toolbar, '{', createBlock(rows), '}');

        container.replaceChildren(panel);

        if (isPreview) {
            return { panel, toolbar };
        }

        document.body.classList.remove('bodyPreviewed');

        view.panel = panel;
        state.mainView = view;

        const canCreateOrBrowse = relationship == null && logicalName && !isCreateMode && pluralName !== 'workflows';

        if (!isCreateMode) {
            toolbar.append(createToolbarButton('View as raw JSON', () => { window.location.href = window.location.href.split('#')[0]; }));
        }

        if (canCreateOrBrowse) {
            const createButton = createToolbarButton('Create new row', async () => {
                setVisible(createButton, false);
                await handleCreateNewRecord(logicalName);
            });
            toolbar.append(createButton);
        }

        if (canCreateOrBrowse && !isSingleColumnValueOnly && !isMultiple) {
            const browseButton = createToolbarButton('View relationships', async () => {
                setVisible(browseButton, false);
                await handleBrowseRelationships(logicalName);
            });
            toolbar.append(browseButton);
        }

        const clientdata = state.originalRecord?.clientdata;
        if (!isMultiple && pluralName === 'workflows' && typeof clientdata === 'string' && clientdata.startsWith('{"')) {
            createMonacoEditorControls(panel, toolbar, singleRecordId, clientdata);
        }

        return { panel, toolbar };
    }

    // ---------------------------------------------------------------------
    // Edit menu
    // ---------------------------------------------------------------------

    function createEditMenu(view, { logicalName, pluralName, id, isCreateMode }) {
        function createCheckbox(label, checked = false) {
            const input = el('input', { type: 'checkbox', checked });
            return { input, row: el('div', { className: 'checkBoxDiv' }, label, input) };
        }

        const bypassPlugins = createCheckbox('Bypass Custom Plugin execution');
        const bypassFlows = createCheckbox('Bypass Power Automate Flow execution');
        const previewChanges = createCheckbox('Preview changes before committing save', true);
        const impersonate = createCheckbox('Impersonate another user');
        const rememberSettings = createCheckbox('Remember these settings (for this environment)');

        const impersonationField = el('select', { on: { change: () => setImpersonationField(impersonationField.value) } },
            el('option', { value: 'systemuserid' }, 'systemuserid'),
            el('option', { value: 'azureactivedirectoryobjectid' }, 'azureactivedirectoryobjectid'));
        const impersonationLabel = el('span', { className: 'impersonationIdFieldLabel' }, 'systemuserid:');
        const impersonationInput = el('input', { placeholder: emptyGuid, on: { input: () => updateImpersonationPreview() } });
        const impersonationPreview = el('span', { className: 'impersonateUserPreview' });

        const impersonateDiv = el('div', { className: 'impersonateDiv pfwa-hidden' },
            el('div', {}, 'Base impersonation on this field: ', impersonationField, '  ', el('i', {}, el('a', { href: impersonationDocsUrl, target: '_blank' }, "What's this?"))),
            el('div', {}, impersonationLabel, impersonationInput, '  ', impersonationPreview));

        const editLink = actionLink('Edit this row', () => editRecord(logicalName, pluralName, id, isCreateMode), 'editLink');
        const cancelLink = actionLink('Cancel', () => reloadPage(pluralName));
        const submitLink = actionLink('Save', async () => {
            setVisible(cancelLink, false);
            setVisible(submitLink, false);
            await submitEdit(pluralName, id, isCreateMode);
        });
        const previewContainer = el('div', { className: 'previewChangesDiv' });
        const savingIndicator = el('div', { className: 'saveInProgressDiv pfwa-hidden' }, 'Saving...');

        const element = el('div', { className: 'editMenuDiv pf-link pfwa-hidden' },
            bypassPlugins.row,
            bypassFlows.row,
            previewChanges.row,
            impersonate.row,
            rememberSettings.row,
            impersonateDiv,
            previewContainer,
            el('div', { className: 'editMenuAction' }, cancelLink),
            el('div', { className: 'editMenuAction' }, submitLink),
            savingIndicator);

        function setImpersonationEnabled(enabled) {
            impersonate.input.checked = enabled;
            setVisible(impersonateDiv, enabled);
        }

        function setImpersonationField(field) {
            impersonationField.value = field;
            impersonationLabel.textContent = field + ':';
            updateImpersonationPreview();
        }

        function setImpersonatedUser(field, userId) {
            setImpersonationEnabled(true);
            impersonationInput.value = userId;
            setImpersonationField(field);
        }

        async function updateImpersonationPreview() {
            if (!impersonationInput.value) {
                impersonationPreview.textContent = '';
                return;
            }

            const result = await odataFetch(apiUrl + `systemusers?$filter=${impersonationField.value} eq '${impersonationInput.value}'&$select=fullname`, false);

            if (result.error) {
                impersonationPreview.textContent = result.error.message;
            } else if (result.value.length == 0) {
                impersonationPreview.textContent = 'user not found';
            } else if (result.value.length == 1) {
                impersonationPreview.textContent = result.value[0].fullname;
            } else {
                impersonationPreview.textContent = 'Something went wrong with retrieving the systemuser.';
            }
        }

        impersonate.input.addEventListener('change', () => setImpersonationEnabled(impersonate.input.checked));

        if (view) {
            view.editMenu = {
                element,
                editLink,
                cancelLink,
                submitLink,
                previewContainer,
                savingIndicator,
                bypassPlugins: bypassPlugins.input,
                bypassFlows: bypassFlows.input,
                previewChanges: previewChanges.input,
                impersonate: impersonate.input,
                impersonationField,
                impersonationInput,
                rememberSettings: rememberSettings.input,
                setImpersonatedUser,
            };
        }

        return [linkSpan(editLink), element];
    }

    function applySettings(menu) {
        let savedSettings = null;

        try {
            savedSettings = JSON.parse(localStorage.getItem(settingsStorageKey));
        } catch {
            // ignore
        }

        if (!savedSettings) {
            return;
        }

        menu.rememberSettings.checked = true;

        if (savedSettings.CallerObjectId) {
            menu.setImpersonatedUser('azureactivedirectoryobjectid', savedSettings.CallerObjectId);
        } else if (savedSettings.MSCRMCallerID) {
            menu.setImpersonatedUser('systemuserid', savedSettings.MSCRMCallerID);
        }

        menu.bypassPlugins.checked = !!savedSettings.bypassCustomPluginExecution;
        menu.bypassFlows.checked = !!savedSettings.bypassFlowExecution;
        menu.previewChanges.checked = !!savedSettings.wasPreviewEnabled;
    }

    function saveSettings(menu, impersonateHeader, bypassCustomPluginExecution, bypassFlowExecution, wasPreviewEnabled) {
        try {
            if (!menu.rememberSettings.checked) {
                localStorage.removeItem(settingsStorageKey);
                return;
            }

            const settings = { ...impersonateHeader, wasPreviewEnabled };

            if (bypassCustomPluginExecution) {
                settings.bypassCustomPluginExecution = true;
            }

            if (bypassFlowExecution) {
                settings.bypassFlowExecution = true;
            }

            localStorage.setItem(settingsStorageKey, JSON.stringify(settings));
        } catch {
            // ignore
        }
    }

    // ---------------------------------------------------------------------
    // Editors
    //
    // An editor is created for every editable column and appended to its slot.
    // Editors have a fieldName, a dataType and a read() method which returns the
    // value to save or throws a UserError. Lookup editors have readChange() instead.
    // ---------------------------------------------------------------------

    function initialValue(fieldName) {
        return state.mode === 'update' ? state.originalRecord[fieldName] : null;
    }

    function createEditor(attribute, slot, optionSets) {
        const attributeType = attribute.AttributeType;
        const isLookup = attributeType === 'Lookup' || attributeType === 'Owner';

        // a lookup editor needs a lookup slot and vice versa
        if (isLookup !== (slot.kind === 'lookup')) {
            return null;
        }

        if (textInputTypes.hasOwnProperty(attributeType)) {
            return createTextEditor(slot, textInputTypes[attributeType]);
        }

        const optionSetOf = (metadataList) => findOptionSet(metadataList, attribute.LogicalName);

        switch (attributeType) {
            case 'Owner':
                return createLookupEditor(slot, ['systemuser', 'team']);
            case 'Lookup':
                // not all lookups are targeted to a table apparently, skip if there are no targets
                return attribute.Targets.length > 0 ? createLookupEditor(slot, attribute.Targets) : null;
            case 'Picklist': {
                const optionSet = optionSetOf(optionSets.picklist);
                return optionSet && createOptionSetEditor(slot, optionSet.Options, { nullable: true, editable: true });
            }
            case 'State': {
                const optionSet = optionSetOf(optionSets.state);
                return optionSet && createOptionSetEditor(slot, optionSet.Options, { nullable: false, editable: false });
            }
            case 'Status': {
                const optionSet = optionSetOf(optionSets.status);
                return optionSet && createOptionSetEditor(slot, optionSet.Options, { nullable: false, editable: true, isStatus: true });
            }
            case 'Boolean': {
                const optionSet = optionSetOf(optionSets.boolean);
                return optionSet && createBooleanEditor(slot, optionSet.FalseOption, optionSet.TrueOption);
            }
            case 'Virtual': {
                const optionSet = attribute.AttributeTypeName?.Value === 'MultiSelectPicklistType' ? optionSetOf(optionSets.multiSelect) : null;
                return optionSet && createMultiSelectEditor(slot, optionSet.Options);
            }
            default:
                return null;
        }
    }

    function createTextEditor(slot, dataType) {
        const { fieldName } = slot;
        const value = initialValue(fieldName);
        const input = el(dataType === 'memo' ? 'textarea' : 'input', { value: value ?? '' });

        // in some very rare cases, dataverse will return an empty string instead of null
        const isEmptyString = value === '';

        if (dataType === 'memo') {
            slot.element.classList.add('memoField');
        }

        slot.element.append(input);

        return {
            fieldName,
            dataType,
            read: () => parseTextInput(fieldName, dataType, input.value, isEmptyString)
        };
    }

    function parseTextInput(fieldName, dataType, rawValue, isEmptyString) {
        switch (dataType) {
            case 'string':
            case 'memo':
            case 'uid':
                return rawValue === '' && !isEmptyString ? null : rawValue;
            case 'int':
                if (rawValue === '') {
                    return null;
                }
                if (!/^-?\d+$/.test(rawValue)) {
                    throw new UserError(`${fieldName} is a whole number. The value ${rawValue} is not compatible.`);
                }
                return parseInt(rawValue);
            case 'decimal':
            case 'money':
            case 'float':
                if (rawValue === '') {
                    return null;
                }
                if (rawValue.includes(',')) {
                    throw new UserError(`${fieldName} is a column of type '${dataType}' and contains a comma (,). Use a dot (.) as the separator.`);
                }
                if (!/^-?[0-9]\d*(\.\d+)?$/.test(rawValue)) {
                    throw new UserError(`${fieldName} is a column of type '${dataType}'. The value ${rawValue} is not compatible.`);
                }
                return parseFloat(rawValue);
            case 'datetime':
                return rawValue === '' ? null : rawValue;
            default:
                throw new UserError(`Unknown data type '${dataType}' for field ${fieldName}. This should not happen.`);
        }
    }

    function createOptionSetEditor(slot, options, { nullable, editable, isStatus = false }) {
        const { fieldName } = slot;
        const value = initialValue(fieldName);
        const select = el('select', { disabled: !editable });

        if (nullable) {
            select.append(el('option', { value: 'null' }, 'null')); // empty option for clearing it
        }

        for (const option of options) {
            const optionElement = el('option', { value: String(option.Value) }, option.Value + ' : ' + option.Label?.UserLocalizedLabel?.Label);

            if (option.State != null) {
                optionElement.dataset.state = String(option.State);
            }

            select.append(optionElement);
        }

        if (value != null && !options.some(o => o.Value === value)) {
            select.append(el('option', { value: String(value) }, value + ' : (unmapped)'));
        }

        if (value != null) {
            select.value = String(value);
        }

        if (isStatus) {
            // keep the statecode in sync with the selected status
            select.addEventListener('change', () => {
                const neededState = select.selectedOptions[0]?.dataset.state;
                const stateSelect = state.mainView.editors.find(e => e.fieldName === 'statecode')?.select;

                if (neededState != null && stateSelect && [...stateSelect.options].some(o => o.value === neededState)) {
                    stateSelect.value = neededState;
                }
            });
        }

        slot.element.append(select);

        return {
            fieldName,
            dataType: 'option',
            select,
            read: () => {
                if (!select.value) {
                    // the select needs to contain a value always, if not, an error happened
                    throw new UserError('there was an error parsing the field ' + fieldName);
                }
                return select.value === 'null' ? null : parseInt(select.value);
            }
        };
    }

    function createBooleanEditor(slot, falseOption, trueOption) {
        const { fieldName } = slot;
        const value = initialValue(fieldName);

        const select = el('select', {},
            el('option', { value: 'null' }, 'null'), // empty option for clearing it
            el('option', { value: 'false' }, 'false : ' + (falseOption?.Label?.UserLocalizedLabel?.Label ?? 'false')),
            el('option', { value: 'true' }, 'true : ' + (trueOption?.Label?.UserLocalizedLabel?.Label ?? 'true')));

        select.value = typeof value === 'boolean' ? String(value) : 'null';

        slot.element.append(select);

        return {
            fieldName,
            dataType: 'bool',
            read: () => {
                switch (select.value) {
                    case 'null': return null;
                    case 'true': return true;
                    case 'false': return false;
                    default: throw new UserError('there was an error parsing the field ' + fieldName);
                }
            }
        };
    }

    function createMultiSelectEditor(slot, options) {
        const { fieldName } = slot;
        const selectedValues = initialValue(fieldName)?.split(',') ?? [];

        // a select which only displays the selected labels, the options are shown in a custom dropdown
        const summaryOption = el('option', { className: 'pfwa-hidden' });
        const fakeSelect = el('select', {}, summaryOption);

        const sortedOptions = [...options].sort((a, b) => a.Value - b.Value);

        const checkboxes = sortedOptions.map(option => ({
            option,
            input: el('input', {
                type: 'checkbox',
                className: 'multiSelectInput',
                checked: selectedValues.includes(option.Value?.toString()),
                on: { change: () => updateSummary() }
            })
        }));

        const dropdown = el('div', { className: 'multiSelectDiv pfwa-hidden' },
            checkboxes.map(({ option, input }) => el('div', { className: 'multiSelectSubDiv' }, input, option.Value + ' : ' + option.Label?.UserLocalizedLabel?.Label)));

        function updateSummary() {
            const labels = checkboxes.filter(c => c.input.checked).map(c => String(c.option.Label?.UserLocalizedLabel?.Label));
            summaryOption.textContent = labels.length > 0 ? labels.join(', ') : '(no options selected)';
        }

        fakeSelect.addEventListener('mousedown', (e) => {
            // prevent the native dropdown from opening
            e.preventDefault();

            setVisible(dropdown, true);
            showOverlay(() => setVisible(dropdown, false));
        });

        updateSummary();

        slot.element.append(el('div', {}, fakeSelect, dropdown));

        return {
            fieldName,
            dataType: 'multiselectoption',
            read: () => {
                const values = checkboxes.filter(c => c.input.checked).map(c => c.option.Value);
                // normalize to null if no options checked
                return values.length > 0 ? values.join(',') : null;
            }
        };
    }

    function createLookupEditor(slot, targets) {
        const { fieldName } = slot;
        const initialTable = slot.lookupType ?? targets[0];

        const tableSelect = el('select', { className: 'lookupSelectTable', disabled: targets.length < 2, on: { change: () => onTableChanged() } },
            targets.map(target => el('option', { value: target }, target)));
        tableSelect.value = initialTable;

        const input = el('input', { placeholder: emptyGuid, on: { input: () => updatePreview() } });
        const preview = el('span', { className: 'lookupPreview' });

        const editPanel = el('div', { className: 'lookupEditMenuDiv pfwa-hidden' },
            el('div', { className: 'lookupEditRow' }, 'Table: ', tableSelect),
            el('div', { className: 'lookupEditRow' }, 'Row id:', input, preview));

        const makeNullNotice = el('div', { className: 'makeLookupNullDiv pfwa-hidden' }, el('b', {}, 'Lookup will be cleared (will be set to null)'));

        const undoItem = linkSpan(actionLink('Undo changes', () => undo()));
        setVisible(undoItem, false);

        const links = el('span', { className: 'lookupEditLinks' },
            linkSpan(actionLink('Edit lookup', () => beginEdit('update'))),
            linkSpan(actionLink('Clear lookup', () => beginEdit('makenull'))),
            undoItem);

        slot.element.append(links, makeNullNotice, editPanel);

        const editor = {
            fieldName,
            dataType: 'lookup',
            editMode: null, // null | 'update' | 'makenull'
            logicalName: initialTable,
            navigationProperty: null,
            readChange
        };

        async function syncSelectedTable() {
            editor.logicalName = tableSelect.value;
            editor.navigationProperty = await retrieveNavigationProperty(tableSelect.value, fieldName);
        }

        async function onTableChanged() {
            await syncSelectedTable();
            await updatePreview();
        }

        async function updatePreview() {
            if (!input.value) {
                preview.textContent = '';
                return;
            }

            if (!tableSelect.value) {
                alert('selectTable does not contain a value. This should not happen.');
                return;
            }

            const entity = await findEntityByLogicalName(tableSelect.value);
            const record = await odataFetch(apiUrl + `${entity.EntitySetName}(${input.value})?$select=${entity.PrimaryNameAttribute}`, false);

            preview.textContent = record.error ? record.error.message : record[entity.PrimaryNameAttribute];
        }

        async function beginEdit(editMode) {
            const navigationProperty = await retrieveNavigationProperty(editor.logicalName, fieldName);

            if (navigationProperty == null) {
                // user was already alerted at this point
                return;
            }

            editor.navigationProperty = navigationProperty;
            editor.editMode = editMode;

            if (editMode === 'makenull') {
                input.value = '';
            }

            setVisible(makeNullNotice, editMode === 'makenull');
            setVisible(editPanel, editMode === 'update');
            setVisible(undoItem, true);
        }

        function undo() {
            editor.editMode = null;
            input.value = '';

            setVisible(undoItem, false);
            setVisible(makeNullNotice, false);
            setVisible(editPanel, false);
        }

        // returns null when the lookup was not touched
        async function readChange() {
            if (editor.editMode !== 'update' && editor.editMode !== 'makenull') {
                return null;
            }

            if (!tableSelect.value) {
                throw new UserError('Error for lookup field: ' + fieldName + '. No table was selected. This should not be possible.');
            }

            const originalValue = slot.id != null ? `/${await retrievePluralName(initialTable)}(${slot.id})` : null;

            if (editor.editMode === 'makenull') {
                return { originalValue, value: null };
            }

            const inputValue = input.value;
            const value = `/${await retrievePluralName(tableSelect.value)}(${inputValue})`;

            if (inputValue === '') {
                throw new UserError('Error for lookup field: ' + fieldName + '. The field was marked for update but it is empty. If you do not want to edit this field, hit "undo changes".');
            }

            if (inputValue.length !== 36) {
                throw new UserError('Error for lookup field: ' + fieldName + '. The value ' + value + ' is not a valid guid.');
            }

            return { originalValue, value };
        }

        // make sure the navigation property matches the selected table
        syncSelectedTable();

        return editor;
    }

    // a transparent layer which closes a custom dropdown when clicking outside of it
    function showOverlay(onClose) {
        const overlay = el('div', {
            id: 'transParentOverlay',
            on: {
                click: () => {
                    overlay.remove();
                    onClose();
                }
            }
        });

        document.body.appendChild(overlay);
    }

    // ---------------------------------------------------------------------
    // Edit, create, delete
    // ---------------------------------------------------------------------

    async function editRecord(logicalName, pluralName, id, isCreateMode) {
        const view = state.mainView;
        state.mode = isCreateMode ? 'create' : 'update';

        setVisible(view.editMenu.editLink, false);

        const attributes = await retrieveEditableAttributes(logicalName, isCreateMode);

        const [picklist, multiSelect, stateOptions, status, boolean] = await Promise.all([
            retrieveOptionSetAttributes(logicalName, 'PicklistAttributeMetadata'),
            retrieveOptionSetAttributes(logicalName, 'MultiSelectPicklistAttributeMetadata'),
            retrieveOptionSetAttributes(logicalName, 'StateAttributeMetadata'),
            retrieveOptionSetAttributes(logicalName, 'StatusAttributeMetadata'),
            retrieveOptionSetAttributes(logicalName, 'BooleanAttributeMetadata'),
        ]);
        const optionSets = { picklist, multiSelect, state: stateOptions, status, boolean };

        const attributesByName = new Map(attributes.map(a => [a.LogicalName, a]));

        // stageid causes many issues for some reason, just leave it out
        attributesByName.delete('stageid');

        // slots are iterated in the order they are displayed, so that changes are listed in the same order
        view.editors = [];

        for (const [fieldName, slot] of view.slots) {
            const attribute = attributesByName.get(fieldName);
            const editor = attribute ? createEditor(attribute, slot, optionSets) : null;

            if (editor) {
                view.editors.push(editor);
            }
        }

        view.panel.classList.add('editing');
        setVisible(view.editMenu.element, true);

        applySettings(view.editMenu);
    }

    // disables all controls while previewing changes, so that no edits can be made after previewing
    function setControlsLocked(view, locked) {
        if (locked) {
            view.lockedControls = [...view.panel.querySelectorAll('input, select, textarea')].filter(c => !c.disabled);
            view.lockedControls.forEach(c => c.disabled = true);
        } else {
            view.lockedControls.forEach(c => c.disabled = false);
            view.lockedControls = [];
        }

        view.panel.classList.toggle('locked', locked);
    }

    function resetSubmitControls() {
        const view = state.mainView;
        const menu = view.editMenu;

        setVisible(menu.savingIndicator, false);
        setVisible(menu.cancelLink, true);
        setVisible(menu.submitLink, true);
        menu.previewContainer.replaceChildren();

        setControlsLocked(view, false);
    }

    function hasChanged(value, originalValue, dataType) {
        if (value === originalValue || (value === '' && originalValue == null)) {
            return false;
        }

        if (dataType === 'memo') {
            return originalValue?.replaceAll('\r\n', '\n') !== value;
        }

        return true;
    }

    // returns a list of { column, payloadKey, oldValue, newValue }
    async function collectChanges(editors, isCreateMode) {
        const changes = [];

        for (const editor of editors) {
            if (editor.dataType === 'lookup') {
                const change = await editor.readChange();

                if (change != null && hasChanged(change.value, change.originalValue, 'lookup')) {
                    changes.push({
                        column: editor.navigationProperty,
                        payloadKey: `${editor.navigationProperty}@odata.bind`,
                        oldValue: change.originalValue,
                        newValue: change.value
                    });
                }
                continue;
            }

            // if not found for some reason, normalize to null to prevent issues
            const originalValue = isCreateMode ? null : (state.originalRecord[editor.fieldName] ?? null);
            const value = editor.read();

            if (hasChanged(value, originalValue, editor.dataType)) {
                changes.push({ column: editor.fieldName, payloadKey: editor.fieldName, oldValue: originalValue, newValue: value });
            }
        }

        return changes;
    }

    function readImpersonationHeader(menu) {
        if (!menu.impersonate.checked) {
            return {};
        }

        const field = menu.impersonationField.value;
        const userId = menu.impersonationInput.value;

        if (userId == null || userId === '') {
            throw new UserError('User impersonation was checked, but ' + field + ' is empty');
        }

        if (userId.length !== 36) {
            throw new UserError('User impersonation input error: ' + userId + ' is not a valid guid.');
        }

        if (field === 'systemuserid') {
            return { MSCRMCallerID: userId };
        } else if (field === 'azureactivedirectoryobjectid') {
            return { CallerObjectId: userId };
        }

        throw new UserError('This should not happen. Wrong value in impersonateAnotherUserSelect: ' + field);
    }

    async function submitEdit(pluralName, id, isCreateMode) {
        const view = state.mainView;
        let changes;
        let impersonateHeader;

        try {
            changes = await collectChanges(view.editors, isCreateMode);
            impersonateHeader = readImpersonationHeader(view.editMenu);
        } catch (error) {
            if (!(error instanceof UserError)) {
                resetSubmitControls();
                throw error;
            }

            alert(error.message);
            resetSubmitControls();
            return;
        }

        if (view.editMenu.previewChanges.checked) {
            previewChanges(changes, pluralName, id, impersonateHeader, isCreateMode);
        } else {
            await commitSave(pluralName, id, changes, impersonateHeader, isCreateMode, false);
        }
    }

    function previewChanges(changes, pluralName, id, impersonateHeader, isCreateMode) {
        const view = state.mainView;

        setControlsLocked(view, true);

        const cancelLink = actionLink('Cancel', () => reloadPage(pluralName));
        const commitLink = actionLink('Commit Save', async () => {
            setVisible(commitLink, false);
            setVisible(cancelLink, false);
            await commitSave(pluralName, id, changes, impersonateHeader, isCreateMode, true);
        });

        view.editMenu.previewContainer.replaceChildren(
            createChangesTable(changes, isCreateMode),
            el('div', { className: 'editMenuAction' }, cancelLink),
            el('div', { className: 'editMenuAction' }, commitLink));
    }

    function createChangesTable(changes, isCreateMode) {
        const table = el('table', { id: 'previewTable' });

        if (changes.length === 0) {
            return table;
        }

        const columns = isCreateMode ? ['column', 'new'] : ['column', 'old', 'new'];
        const cellText = (value) => value == null ? '(empty)' : String(value);

        table.append(
            el('thead', {}, el('tr', {}, columns.map(column => el('td', {}, column)))),
            el('tbody', {}, changes.map(change => el('tr', {},
                el('td', {}, change.column),
                isCreateMode ? null : el('td', {}, cellText(change.oldValue)),
                el('td', {}, cellText(change.newValue))))));

        return table;
    }

    async function commitSave(pluralName, id, changes, impersonateHeader, isCreateMode, wasPreviewEnabled) {
        const menu = state.mainView.editMenu;
        const bypassCustomPluginExecution = menu.bypassPlugins.checked;
        const bypassFlowExecution = menu.bypassFlows.checked;

        const headers = {
            ...odataJsonHeaders,
            ...(isCreateMode ? { 'Prefer': 'return=representation' } : { 'If-Match': '*' }),
            ...impersonateHeader
        };

        if (bypassCustomPluginExecution) {
            headers['MSCRM.BypassCustomPluginExecution'] = true;
        }

        if (bypassFlowExecution) {
            headers['MSCRM.SuppressCallbackRegistrationExpanderJob'] = true;
        }

        saveSettings(menu, impersonateHeader, bypassCustomPluginExecution, bypassFlowExecution, wasPreviewEnabled);

        setVisible(menu.savingIndicator, true);

        const response = await fetch(isCreateMode ? apiUrl + pluralName : apiUrl + pluralName + '(' + id + ')', {
            method: isCreateMode ? 'POST' : 'PATCH',
            headers,
            body: JSON.stringify(Object.fromEntries(changes.map(change => [change.payloadKey, change.newValue])))
        });

        if (!response.ok) {
            await alertErrorResponse(response);
            resetSubmitControls();
            return;
        }

        if (isCreateMode) {
            const json = await response.json();
            const { primaryIdAttribute } = await retrieveEntityInfo(pluralName);
            window.location.href = apiUrl + pluralName + '(' + json[primaryIdAttribute] + ')#p';
        } else {
            reloadPage(pluralName);
        }
    }

    async function handleCreateNewRecord(logicalName) {
        const creatableAttributes = await retrieveAttributes(logicalName, 'IsValidForCreate');
        const emptyRecord = {};

        for (const attribute of creatableAttributes) {
            const isLookup = (attribute.AttributeType === 'Lookup' && attribute.Targets.length > 0) || attribute.AttributeType === 'Owner';
            emptyRecord[isLookup ? '_' + attribute.LogicalName + '_value' : attribute.LogicalName] = null;
        }

        // 'manually' add this in, because we want to be able to see this
        // it's not listed as creatable, but it actually is
        emptyRecord['statecode'] = null;

        await renderPanel(emptyRecord, document.body, state.currentPluralName, { isCreateMode: true });
        await editRecord(logicalName, state.currentPluralName, null, true);
    }

    async function deleteRecord(pluralName, id, deleteLink) {
        setVisible(deleteLink, false);

        try {
            if (!confirm('Are you sure you want to delete this row?')) {
                return;
            }

            const response = await fetch(apiUrl + pluralName + '(' + id + ')', { method: 'DELETE' });

            if (response.status === 204) {
                if (confirm('Row was deleted. Click OK to close this page.')) {
                    window.close();
                }
            } else {
                const json = await response.json();

                if (json.error) {
                    alert(json.error.message);
                }
            }
        } finally {
            setVisible(deleteLink, true);
        }
    }

    async function disassociate(sendDisassociateRequest) {
        if (!confirm('Are you sure you want to disassociate this row?')) {
            return;
        }

        if (await sendDisassociateRequest()) {
            alert('Row was disassociated. This page will now reload.');
            window.location.reload();
        }
    }

    function disassociateRowManyToOne(navigationProperty) {
        return disassociate(() => {
            const recordId = location.pathname.split(')')[0].split('(')[1];
            return sendRequest(apiUrl + state.originalPluralName + '(' + recordId + ')', 'PATCH', { [navigationProperty + '@odata.bind']: null });
        });
    }

    function disassociateRowOneToMany(pluralName, id, navigationProperty) {
        return disassociate(() => sendRequest(apiUrl + pluralName + '(' + id + ')', 'PATCH', { [navigationProperty + '@odata.bind']: null }));
    }

    function disassociateRowManyToMany(pluralName, id) {
        return disassociate(() => sendRequest(window.location.pathname + '/$ref?$id=' + location.origin + apiUrl + pluralName + '(' + id + ')', 'DELETE'));
    }

    async function reloadPage(pluralName) {
        if (pluralName === 'workflows') {
            window.location.reload(); // the monaco editor fails us when re-initializing, so just reload
        } else {
            await makeItPretty();
        }
    }

    // ---------------------------------------------------------------------
    // Preview
    // ---------------------------------------------------------------------

    async function previewRecord(pluralName, id) {
        const mainPanel = state.mainView.panel;
        mainPanel.classList.add('prePreviewed');
        document.body.classList.add('bodyPreviewed');

        const previewPanel = el('div', { className: 'panel previewPanel prePreviewed' });
        document.body.appendChild(previewPanel);

        const response = await odataFetch(apiUrl + pluralName + '(' + id + ')', false);

        const { toolbar } = await renderPanel(response, previewPanel, pluralName, { isPreview: true });

        toolbar.append(createToolbarButton('Close Preview', () => {
            if (document.getElementsByClassName('previewPanel').length === 1) {
                mainPanel.classList.remove('prePreviewed');
                document.body.classList.remove('bodyPreviewed');
            }

            previewPanel.remove();
        }));

        previewPanel.scrollIntoView();
    }

    // ---------------------------------------------------------------------
    // Relationship browser
    // ---------------------------------------------------------------------

    async function handleBrowseRelationships(logicalName) {
        const relationships = await retrieveRelationships();
        if (relationships == null) {
            alert(`An error occurred while retrieving relationships. Cannot continue.`);
            return;
        }

        const relationshipsDiv = el('div', { className: 'relationshipList' });
        const filterInput = el('input', {
            type: 'text',
            on: { input: () => renderRelationships(relationshipsDiv, relationships, filterInput.value, logicalName) }
        });

        document.body.replaceChildren(el('pre', { id: 'relationshipBrowser' },
            el('button', { className: 'closeRelationshipsButton', on: { click: () => window.location.reload() } }, 'Close relationships'),
            el('h3', {}, 'Filter relationships'),
            filterInput,
            relationshipsDiv));

        renderRelationships(relationshipsDiv, relationships, null, logicalName);
    }

    function filterRelationships(relationships, fields, filter) {
        return relationships.filter(r => fields.some(field => r[field]?.toLowerCase()?.includes(filter)));
    }

    function renderRelationships(container, relationships, filter, logicalName) {
        let manyToMany = relationships.ManyToManyRelationships;
        let oneToMany = relationships.OneToManyRelationships;
        let manyToOne = relationships.ManyToOneRelationships;

        if (filter) {
            const lowerCaseFilter = filter.toLowerCase();

            manyToMany = filterRelationships(manyToMany, relationshipFilterFields.manyToMany, lowerCaseFilter);
            oneToMany = filterRelationships(oneToMany, relationshipFilterFields.oneToMany, lowerCaseFilter);
            manyToOne = filterRelationships(manyToOne, relationshipFilterFields.oneToMany, lowerCaseFilter);
        }

        container.replaceChildren();
        appendChildren(container, [
            el('h2', {}, 'Many to Many Relationships'),
            manyToMany.map(r => createManyToManyRelationshipLink(r, logicalName)),
            el('h2', {}, 'Many to One Relationships (lookups)'),
            manyToOne.map(r => createManyToOneRelationshipLink(r, logicalName)),
            el('h2', {}, 'One to Many Relationships'),
            oneToMany.map(r => createOneToManyRelationshipLink(r, logicalName))
        ]);
    }

    function createManyToManyRelationshipLink(relationship, logicalName) {
        if (relationship.Entity1LogicalName === logicalName) {
            return createRelationshipLink(relationship.Entity1NavigationPropertyName);
        } else if (relationship.Entity2LogicalName === logicalName) {
            return createRelationshipLink(relationship.Entity2NavigationPropertyName);
        }

        alert(`Could find the navigation property for many to many relationship ${relationship.SchemaName}`);
        return null;
    }

    function createManyToOneRelationshipLink(relationship, logicalName) {
        if (relationship.ReferencingEntity !== logicalName) {
            alert(`Could find the navigation property for many to one relationship '${relationship.SchemaName}'`);
            return null;
        }

        const navigationProperty = relationship.ReferencingEntityNavigationPropertyName;
        const link = createRelationshipLink(navigationProperty);

        // there is nothing to browse when the lookup is empty
        if (state.originalRecord['_' + navigationProperty + '_value'] == null && !state.foundNavigationProperties.includes(navigationProperty)) {
            link.removeAttribute('href');
            link.classList.add('disabledLink');
        }

        return link;
    }

    function createOneToManyRelationshipLink(relationship, logicalName) {
        if (relationship.ReferencedEntity === logicalName) {
            return createRelationshipLink(relationship.ReferencedEntityNavigationPropertyName);
        }

        alert(`Could find the navigation property for one to many relationship '${relationship.SchemaName}'`);
        return null;
    }

    function createRelationshipLink(navigationProperty) {
        return el('a', { className: 'relationshipLink', target: '_blank', href: window.location.pathname + '/' + navigationProperty + '#pr' }, navigationProperty);
    }

    // ---------------------------------------------------------------------
    // Flow editor (monaco)
    // ---------------------------------------------------------------------

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

    function loadPageScript(path, onLoad) {
        document.head.appendChild(el('script', { type: 'text/javascript', src: chrome.runtime.getURL(path), on: onLoad ? { load: onLoad } : {} }));
    }

    function createMonacoEditorControls(mainPanel, toolbar, recordId, clientdata) {
        if (checkIfJsonViewerEnabled()) {
            toolbar.append(el('button', { type: 'button', disabled: true }, 'Currently, editing Flows is only possible in Chrome. Work in progress.'));
            return;
        }

        // the editor runs in the page context, pass the data it needs through the dataset
        Object.assign(mainPanel.dataset, {
            chromeRuntimeUrl: chrome.runtime.getURL(''),
            flowName: state.originalRecord.name,
            recordId,
            apiUrl,
            clientdata
        });

        loadPageScript('libs/monaco/loader.js', () => loadPageScript('libs/monaco/initMonaco.js'));
    }

    // ---------------------------------------------------------------------
    // Page
    // ---------------------------------------------------------------------

    async function makeItPretty() {
        addMainCss();

        if (window.location.hash === '#pf') {
            document.body.replaceChildren('Loading your flow...');
        }

        const jsonObject = await odataFetch(window.location.href, false);

        state.currentPluralName = window.location.pathname.split('(').shift().split('/').pop();

        await renderPanel(jsonObject, document.body, state.currentPluralName);

        // do this for Edge, who overrides our styles
        for (const panel of document.getElementsByClassName('panel')) {
            panel.style.display = 'block';
        }
    }

    function addMainCss() {
        if (document.getElementById('pfwaStyles')) {
            return;
        }

        // remove the styles of the browser's json viewer
        for (const style of [...document.head.getElementsByTagName('style')]) {
            style.remove();
        }

        document.head.appendChild(el('style', { id: 'pfwaStyles', type: 'text/css' }, `
            body {
                font-size: 13px !important;
            }

            button, select, textarea, input {
                font-family: Arial !important;
                font-size: 13px !important;
            }

            .pfwa-hidden {
                display: none !important;
            }

            .pfwa-root {
                margin: 0;
            }

            .pfwa-block {
                padding-left: 3ch;
            }

            .pfwa-toolbar {
                position: absolute;
                top: 20px;
                right: 28px;
                display: flex;
                flex-direction: row-reverse;
                gap: 18px;
            }

            .pfwa-toolbar button {
                height: 30px;
                min-width: 122px;
                padding: 0 4px;
                cursor: pointer;
            }

            .pfwa-toolbar button:disabled {
                cursor: default;
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
                display:grid;
                z-index: 100;
                height: fit-content;
            }

            #relationshipBrowser input {
                max-width: 500px;
            }

            .closeRelationshipsButton {
                position: fixed;
                height: 30px;
                top: 20px;
                right: 20px;
            }

            .relationshipList {
                display: grid;
            }

            .relationshipLink {
                margin: 4px 0px;
                width: fit-content;
            }

            .disabledLink {
                color: LinkText;
                text-decoration: none;
            }

            @media (prefers-color-scheme: dark) {
                #relationshipBrowser {
                    background-color: #28282B;
                }

                .disabledLink {
                    color: #E9E9E9;
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

            .panel option:empty {
                display:none;
            }

            /* fields */

            .field {
                display: inline-flex;
                cursor: pointer;
            }

            .panel .field:not(.lookupField), .panel .lookupDisplay {
                margin-right: 24px;
                padding-right: 16px;
            }

            .copyButton {
                color: dimgray;
                display: none;
                cursor: pointer;
            }

            @media (prefers-color-scheme: dark) {
                .copyButton {
                    color: darkgray;
                }
            }

            .panel:not(.editing) .field:hover .copyButton {
                display: unset;
            }

            .copyIcon, .copiedIcon {
                width: 16px;
                position: absolute;
                padding-left: 4px;
            }

            .copiedNotification {
                display: none;
            }

            .field.copied .copiedNotification {
                display: unset;
            }

            .field.copied .copyIcon {
                display: none;
            }

            .copiedText {
                padding-left: 24px;
            }

            /* edit mode */

            .lookupEdit {
                display: none;
            }

            .editing .lookupEdit {
                display: unset;
            }

            .editing .lookupDisplay, .editing .deleteRow, .locked .lookupEditLinks {
                display: none;
            }

            .editing .field {
                cursor: auto;
                height: 20px;
            }

            .editing .field.memoField {
                display: inline;
            }

            .editing .field.memoField textarea {
                display: block;
            }

            .editing .lookupField {
                margin: 1px 0 -2px 0;
            }

            .checkBoxDiv {
                display: flex;
                padding: 1px 0 1px 1ch;
                align-items: center;
            }

            .panel .checkBoxDiv input {
                width: 25px;
            }

            .impersonateDiv > div, .lookupEditRow, .makeLookupNullDiv {
                padding-left: 3ch;
            }

            .editMenuAction, .saveInProgressDiv {
                padding-left: 1ch;
            }

            .lookupEditLinks .pf-link {
                margin-left: 3ch;
            }

            .lookupSelectTable {
                height: 22px;
            }

            .lookupPreview {
                margin-left: 10px;
            }

            /* changes preview table */

            .panel table {
                color: black;
                margin: 1.2em 0 1.2em 26px;
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

            /* panels */

            .panel {
                padding:10px;
                margin:0px;
                text-wrap:nowrap;
                display:block;
            }

            .mainPanel, .previewPanel {
                position: relative;
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

            /* flow editor */

            .monacoContainer {
                height: 94vh;
                width: 98vw;
                box-sizing: border-box;
            }

            .monacoActions {
                height: 2em;
                display: flex;
                align-items: center;
                gap: 8px;
                border-top: 1px solid #aaa;
                padding: 0.2em;
                box-sizing: border-box;
            }

            .monacoActions button, .monacoActions input, .monacoLabel {
                cursor: pointer;
            }

            .monacoActions input {
                width: unset;
            }
            `));
    }

    await makeItPretty();
})()
