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
    const odataCountKey = '@odata.count';
    const nextLinkKey = '@odata.nextLink';

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
    // Options: className, dataset, style, on (event handlers), and any other element property.
    function el(tag, options = {}, ...children) {
        const { className, dataset, style, on, ...properties } = options;
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

    // An action which runs onClick when clicked. While onClick runs, the action is busy: it stays in place,
    // greyed out, and ignores clicks. Actions never hide themselves, so a double click can't land on
    // another action that moved under the cursor. An action can also be disabled for longer, e.g. while saving.
    function actionLink(text, onClick, className) {
        const link = el('a', {
            className,
            href: '#',
            on: {
                click: async (e) => {
                    e.preventDefault();

                    if (link.classList.contains('busy') || link.classList.contains('disabled')) {
                        return;
                    }

                    link.classList.add('busy');

                    try {
                        await onClick(e);
                    } finally {
                        link.classList.remove('busy');
                    }
                }
            }
        }, text);

        return link;
    }

    /*
     * Layout rule: an action sits on the line of the thing it acts on.
     * - actions on a row (or lookup) are the first line inside its braces: '{' / 'new tab · form' / 'key: value' / '}'
     * - actions on the page are in the header, the sticky line at the top of a panel
     * The data itself is only 'key: value'.
     */

    // secondary text, e.g. the row count in the header
    function createMuted(text) {
        return el('span', { className: 'muted' }, text);
    }

    // the actions on a row, as the first line inside its braces, e.g. 'new tab · form · preview'
    function createActionRow(actions, extraClass) {
        const visibleActions = actions.filter(Boolean);
        return visibleActions.length === 0 ? null : el('div', { className: ['pfwa-row', 'actionRow', 'actions', extraClass].filter(Boolean).join(' ') }, visibleActions);
    }

    // actions after a value on the same line, e.g. a solutionid's 'new tab'
    function createInlineActions(actions) {
        return el('span', { className: 'inlineActions actions' }, actions);
    }

    // The sticky line at the top of a panel: what is shown on the left, the actions on the page on the right.
    // The read parts are shown normally, the edit parts while editing (see .readOnly and .editOnly).
    // While editing, the edit options and the changes to commit are shown below the line.
    function createHeader() {
        const header = {
            readContext: el('span', { className: 'headerContext readOnly' }),
            editContext: el('span', { className: 'headerContext editOnly' }),
            readActions: el('span', { className: 'headerActions actions readOnly' }),
            editActions: el('span', { className: 'headerActions actions editOnly' }),
            editBody: el('div', { className: 'headerBody editOnly' }), // below the line: the edit options and the changes to commit
        };

        header.element = el('div', { className: 'pfwa-toolbar' }, header.readContext, header.editContext, header.readActions, header.editActions, header.editBody);

        return header;
    }

    function setActions(container, actions) {
        container.replaceChildren(...actions.filter(Boolean));
    }

    // a row of the json-like output, e.g. 'key: value'. Pass null as key for array items.
    // The key can be a string or a node, e.g. a copyable key.
    function createRow(key, ...content) {
        const keyNodes = key == null ? null : [typeof key === 'string' ? el('span', { className: 'keySpan' }, key) : key, ': '];
        return el('div', { className: 'pfwa-row' }, keyNodes, content);
    }

    // an indented block of rows, e.g. the contents of an object
    function createBlock(rows) {
        return el('div', { className: 'pfwa-block' }, rows);
    }


    // ---------------------------------------------------------------------
    // Web Api
    // ---------------------------------------------------------------------

    // GET request for json. With memoize, a successful response is remembered for as long as the page is open
    async function odataFetch(url, memoize) {
        if (memoize && memoizedCalls.hasOwnProperty(url)) {
            return memoizedCalls[url];
        }

        const response = await fetch(url, { headers: { 'Prefer': 'odata.include-annotations="*"', 'Cache-Control': 'no-cache' } });

        const responseJson = await response.json();

        // a failed response is not remembered, so the request is retried the next time
        if (memoize && response.ok) {
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

    // Table definitions are loaded lazily: a filtered query per table when its primary attributes are needed,
    // and the list of all table names only when a name mapping is needed that is not known yet.

    const retrievedEntityDefinitions = []; // definitions retrieved by the filtered queries

    // filterProperty: LogicalName or EntitySetName
    async function retrieveEntityDefinition(filterProperty, value) {
        const json = await odataFetch(`${apiUrl}EntityDefinitions?$select=LogicalName,EntitySetName,PrimaryIdAttribute,PrimaryNameAttribute&$filter=(${filterProperty} eq '${value}')`, true);
        const entity = json.value?.[0];

        if (entity && !retrievedEntityDefinitions.includes(entity)) {
            retrievedEntityDefinitions.push(entity);
        }

        return entity;
    }

    function findEntityByLogicalName(logicalName) {
        return retrieveEntityDefinition('LogicalName', logicalName);
    }

    function findEntityByPluralName(pluralName) {
        return retrieveEntityDefinition('EntitySetName', pluralName);
    }

    // LogicalName and EntitySetName of all tables
    async function retrieveAllTableNames() {
        const json = await odataFetch(apiUrl + 'EntityDefinitions?$select=EntitySetName,LogicalName', true);

        return json.value ?? [];
    }

    // finds the table by one of its names, without a request when its definition was retrieved before
    async function findTableName(nameProperty, value) {
        return retrievedEntityDefinitions.find(e => e[nameProperty] === value)
            ?? (await retrieveAllTableNames()).find(e => e[nameProperty] === value);
    }

    async function retrievePluralName(logicalName) {
        return (await findTableName('LogicalName', logicalName))?.EntitySetName;
    }

    async function retrieveLogicalName(pluralName) {
        return (await findTableName('EntitySetName', pluralName))?.LogicalName;
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

    // the rows of a metadata response. Throws the error of Dataverse when the request failed, e.g. a missing privilege
    function readMetadataRows(json) {
        if (json.error) {
            throw new Error(json.error.message);
        }

        return json.value;
    }

    // validityProperty: IsValidForUpdate, IsValidForCreate
    async function retrieveAttributes(logicalName, validityProperty) {
        const json = await odataFetch(`${apiUrl}EntityDefinitions(LogicalName='${logicalName}')/Attributes?$filter=${validityProperty} eq true`, true);

        return readMetadataRows(json);
    }

    // metadataType: PicklistAttributeMetadata, MultiSelectPicklistAttributeMetadata, StateAttributeMetadata, etc.
    async function retrieveOptionSetAttributes(logicalName, metadataType) {
        const json = await odataFetch(`${apiUrl}EntityDefinitions(LogicalName='${logicalName}')/Attributes/Microsoft.Dynamics.CRM.${metadataType}?$select=LogicalName&$expand=OptionSet,GlobalOptionSet`, true);

        return readMetadataRows(json);
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

    // Action labels are short and lowercase, the tooltip explains what they do

    // opens the row in the Web Api in a new tab
    function createWebApiLink(guid, pluralName) {
        return el('a', { target: '_blank', title: 'Open this row in the Web Api in a new tab', href: `${apiUrl}${pluralName}(${stripBraces(guid)})#p` }, 'new tab');
    }

    function createFormLink(logicalName, guid) {
        return el('a', { target: '_blank', title: 'Open this row in a form in a new tab', href: `/main.aspx?etn=${logicalName}&id=${guid}&pagetype=entityrecord` }, 'form');
    }

    // the actions on a row of a collection or a preview, e.g. 'new tab · form · disassociate'
    function createRowActions(logicalName, pluralName, recordId, relationship) {
        if (!logicalName) {
            return [];
        }

        if (recordId == null || recordId === '') {
            return [createMuted('could not generate links')];
        }

        return [
            createWebApiLink(recordId, pluralName),
            createFormLink(logicalName, recordId),
            relationship ? createDisassociateLink(relationship, pluralName, recordId) : null
        ];
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

        const link = actionLink('disassociate', () => disassociate(), 'dangerLink');
        link.title = 'Remove the relationship between this row and the parent row';

        return link;
    }

    // ---------------------------------------------------------------------
    // Fields
    // ---------------------------------------------------------------------

    // A tooltip for everything that can be copied (an element with a copyText), showing exactly what will be copied.
    // One listener for the whole page sets it when hovering, because there is nothing to copy while editing.
    function showCopyTooltip(event) {
        const element = event.target.closest('.field, .copyableKey');

        if (element?.copyText == null) {
            return;
        }

        const maxLength = 100;
        const text = element.copyText;

        element.title = element.closest('.editing') ? '' : 'Click to copy: ' + (text.length > maxLength ? text.slice(0, maxLength) + '…' : text);
    }

    /*
     * A value which can be copied to the clipboard by clicking it. Only the raw value is copied, not the display text.
     * The whole field is clickable, but only the part which is copied is underlined on hover.
     * Options:
     *   label: display only text in front of the value, e.g. 'Name: '
     *   formattedValue: display only text after the value, shown dimmed
     *   extraClasses: additional css classes
     * A very long value, e.g. the description of an email, is cut off at the edge of the page with '…', followed by 'show all'
     * (not in the header, which cuts off by itself). Clicking the field still copies the full value.
     */
    function createField(cls, value, { label, formattedValue, extraClasses = [] } = {}) {
        const text = String(value);
        const displayText = formatDisplayText(value);
        const isLong = displayText.length > longValueLength && !extraClasses.includes('inlineField');
        const formattedPart = formattedValue == null ? null : el('span', { className: 'formattedValue' }, ' : ' + formatDisplayText(formattedValue));

        const field = el('span', { className: [cls, 'field', ...extraClasses, isLong ? 'longValue' : null].filter(Boolean).join(' ') },
            label,
            el('span', { className: 'copyValue' }, displayText),
            formattedPart);

        if (isLong) {
            field.append(createShowAllLink(field, displayText.length));
        }

        field.copyText = text;
        field.addEventListener('click', () => copyFieldToClipboard(field, text));

        return field;
    }

    // values with more characters than this are cut off, see .longValue
    const longValueLength = 200;

    // shows all of a value which was cut off, wrapped to the width of the page, or cuts it off again
    function createShowAllLink(field, length) {
        const link = el('a', {
            className: 'showAllLink',
            href: '#',
            title: `Show all ${length} characters`,
            on: {
                click: (e) => {
                    e.preventDefault();
                    e.stopPropagation(); // not a click on the field, which copies it

                    const showAll = field.classList.toggle('showingAll');

                    link.textContent = showAll ? 'show less' : 'show all';
                    link.title = showAll ? 'Cut it off again' : `Show all ${length} characters`;
                }
            }
        }, 'show all');

        return link;
    }

    // a column name which copies itself to the clipboard when clicked
    function createCopyableKey(key) {
        const keySpan = el('span', { className: 'keySpan copyableKey' }, key);
        keySpan.copyText = key;

        keySpan.addEventListener('click', async () => {
            // no copying while editing, same as fields
            if (keySpan.closest('.editing')) {
                return;
            }

            try {
                await navigator.clipboard.writeText(key);
                keySpan.classList.add('copied');
                setTimeout(() => keySpan.classList.remove('copied'), 400);
            } catch {
                alert('Failed to copy');
            }
        });

        return keySpan;
    }

    // a field which gets an input appended when the row is edited
    function createEditableField(cls, value, fieldName, view, fieldOptions) {
        const field = createField(cls, value, fieldOptions);

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
     *   relationship: the relationship definition when browsing related rows
     *   view: the main view to register editable fields in, null when not editable
     * The actions on a row are rendered by the caller, as the first line inside its braces.
     */
    const nestedRenderOptions = {};

    // the 'key: value' rows of a row: nested objects and arrays are listed first, followed by the columns
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
                columnRows.push(createRow(createCopyableKey(key), await renderPropertyValue(record, key, options, consumedKeys)));
            }
        }

        return [...nestedRows, ...columnRows];
    }

    async function renderNestedValue(value) {
        if (!Array.isArray(value)) {
            return ['{', createBlock(await renderPropertyRows(value, nestedRenderOptions)), '}'];
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
            return [createField(cls, value, { extraClasses: ['lookupField'] }), createLookupSlot(view, key, null, null)];
        }

        const formattedValue = record[key + formattedValueType];
        if (formattedValue) {
            consumedKeys.add(key + formattedValueType);

            // toString the value because it can be a number. The formattedValue is always a string
            const isDifferent = value?.toString() !== formattedValue;

            return createEditableField(cls, value, key, view, isDifferent ? { formattedValue } : undefined);
        }

        if (options.logicalName !== 'solution' && key === 'solutionid' && value != null) {
            return [createEditableField(cls, value, key, view), createInlineActions([createWebApiLink(value, 'solutions')])];
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

        const previewLink = actionLink('preview', () => previewRecord(pluralName, id));
        previewLink.title = 'Show this row next to the current one';

        const lookupDisplay = el('span', { className: 'lookupDisplay' },
            '{',
            createBlock([
                createActionRow([createWebApiLink(id, pluralName), createFormLink(lookupLogicalName, guid), previewLink]),
                createRow(null, createField('guid', id)),
                createRow(null, createField(determineType(formattedValue), formattedValue, { label: 'Name: ' })),
                createRow(null, createField(determineType(lookupLogicalName), lookupLogicalName, { label: 'LogicalName: ' })),
                createRow(null, createField(determineType(navigationProperty), navigationProperty, { label: 'NavigationProperty: ' })),
            ]),
            '}');

        // in edit mode, only the name is displayed
        const lookupName = el('span', { className: 'lookupEdit' }, createField('string', formattedValue, { extraClasses: ['lookupField'] }));

        return [lookupDisplay, lookupName, createLookupSlot(view, key, guid, lookupLogicalName)];
    }

    // Collections: the remaining top level keys (e.g. @odata.nextLink) followed by the rows, each with its actions as its first line
    async function renderCollectionRows(json, options) {
        const { logicalName, pluralName, primaryIdAttribute, relationship } = options;
        const rows = [];

        for (const key of Object.keys(json)) {
            if (key === 'value' || hiddenCollectionKeys.includes(key)) {
                continue;
            }

            const value = json[key];
            rows.push(createRow(key, isObject(value) ? await renderNestedValue(value) : createField(determineType(value), value)));
        }

        const records = [];
        for (const record of json.value) {
            const actions = createRowActions(logicalName, pluralName, record[primaryIdAttribute], relationship);
            records.push(createRow(null, '{', createBlock([createActionRow(actions), ...await renderPropertyRows(record, options)]), '}'));
        }

        rows.push(createRow('value', '[', createBlock(records), ']'));

        return rows;
    }

    // the next page opens in the same tab, like paging usually does
    function createNextPageLink(nextLink) {
        return el('a', { href: nextLink + '#p', title: 'Open the next page of rows' }, 'next page');
    }

    // e.g. '50 rows', '50 of 1234 rows' or '50 of 5000+ rows'
    // The total count is only known when it is requested with $count=true, it is -1 otherwise
    function describeRowCount(json) {
        const rowCount = json.value.length;
        const totalCount = isNumber(json[odataCountKey]) ? json[odataCountKey] : json[totalRecordCountKey];
        const isLimitExceeded = json[totalRecordCountLimitExceededKey] === true;

        if (!isNumber(totalCount) || totalCount < 0 || (totalCount === rowCount && !isLimitExceeded)) {
            return `${rowCount} rows`;
        }

        return `${rowCount} of ${totalCount}${isLimitExceeded ? '+' : ''} rows`;
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

    // Renders a Web Api response into the container: the header, followed by the response. Returns the rendered panel and its header.
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
            relationship,
            view: null,
        };

        const header = createHeader();
        const canCreateOrBrowse = !isPreview && relationship == null && logicalName && !isCreateMode;

        let rows;
        let rowActions = []; // the actions on the main row, its first line
        let singleRecordId = null;

        if (!logicalName) {
            // sdk messages other than retrieve or retrievemultiple
            header.readContext.append(initialPluralName);
            rows = await renderPropertyRows(json, { ...recordOptions, relationship: null });
        } else if (isMultiple) {
            const title = relationship != null ? 'related ' + pluralName : pluralName;
            setTitle(title);
            header.readContext.append(title, ' ', createMuted(describeRowCount(json)));
            rows = await renderCollectionRows(json, recordOptions);
        } else {
            setTitle(logicalName);
            singleRecordId = json[entity.primaryIdAttribute];
            const hasRecordId = singleRecordId != null && singleRecordId !== '';

            rows = await renderPropertyRows(json, { ...recordOptions, view });

            if (isCreateMode) {
                header.readContext.append('new ' + logicalName);
            } else {
                header.readContext.append(isPreview ? 'preview ' + logicalName : logicalName);

                // the primary name in bold, e.g. 'contact · Amie Gonzales'. Only the table name (and no dot) when the row has no name
                const primaryName = json[entity.primaryNameAttribute];

                if (primaryName != null && primaryName !== '') {
                    header.readContext.append(el('span', { className: 'separator' }, '·'), el('b', {}, createField(determineType(primaryName), primaryName, { extraClasses: ['inlineField'] })));
                } else if (isSingleColumnValueOnly) {
                    header.readContext.append(' ', createMuted(location.pathname.split('/').pop()));
                }
            }

            if (isPreview) {
                rowActions = createRowActions(logicalName, pluralName, singleRecordId, null);
            } else if (isCreateMode || hasRecordId) {
                const editPluralName = await retrievePluralName(logicalName);
                const id = isCreateMode ? null : stripBraces(singleRecordId);
                const menu = createEditMenu(view, { logicalName, pluralName: editPluralName, id, isCreateMode, header });

                if (!isCreateMode) {
                    rowActions = [
                        createFormLink(logicalName, singleRecordId),
                        menu.editLink,
                        canCreateOrBrowse && !isSingleColumnValueOnly ? createRelationshipsLink(logicalName) : null,
                        relationship ? createDisassociateLink(relationship, pluralName, singleRecordId) : null,
                        createDeleteLink(editPluralName, id),
                    ];
                }
            }
        }

        const panel = el('pre', { className: isPreview ? 'pfwa-root' : 'pfwa-root mainPanel panel' },
            header.element,
            '{',
            createBlock([createActionRow(rowActions, 'readOnly'), ...rows]),
            '}');

        container.replaceChildren(panel);

        if (isPreview) {
            return { panel, header };
        }

        document.body.classList.remove('bodyPreviewed');

        view.panel = panel;
        state.mainView = view;

        const newRowLink = canCreateOrBrowse ? actionLink('new row', () => handleCreateNewRecord(logicalName)) : null;

        if (newRowLink) {
            newRowLink.title = 'Create a new row in this table';
        }

        setActions(header.readActions, [
            isMultiple && typeof json[nextLinkKey] === 'string' ? createNextPageLink(json[nextLinkKey]) : null,
            newRowLink,
            isCreateMode ? null : el('a', { href: window.location.href.split('#')[0], title: 'Show the response of the Web Api as is' }, 'raw json'),
        ]);

        const clientdata = state.originalRecord?.clientdata;
        if (!isMultiple && pluralName === 'workflows' && typeof clientdata === 'string' && clientdata.startsWith('{"')) {
            createMonacoEditorControls(panel, header, singleRecordId, clientdata);
        }

        return { panel, header };
    }

    function createRelationshipsLink(logicalName) {
        const link = actionLink('relationships', () => handleBrowseRelationships(logicalName));
        link.title = 'Browse the relationships of this row';

        return link;
    }

    function createDeleteLink(pluralName, id) {
        const link = actionLink('delete', () => deleteRecord(pluralName, id), 'dangerLink');
        link.title = 'Delete this row';

        return link;
    }

    // ---------------------------------------------------------------------
    // Edit menu
    // ---------------------------------------------------------------------

    /*
     * The edit controls of the main row:
     *   editLink: the 'edit' action, on the first line of the row
     *   in the header while editing: what is edited, the number of changes, save and cancel,
     *   and below that the options, the impersonation inputs and the changes to commit
     */
    function createEditMenu(view, { logicalName, pluralName, id, isCreateMode, header }) {
        function createCheckbox(label, title, checked = false) {
            const input = el('input', { type: 'checkbox', checked });
            return { input, label: el('label', { className: 'editOption', title }, input, label) };
        }

        const previewChanges = createCheckbox('preview changes', 'Show the changes before committing the save', true);
        const bypassPlugins = createCheckbox('bypass plugins', 'Bypass custom plugin execution');
        const bypassFlows = createCheckbox('bypass flows', 'Bypass Power Automate flow execution');
        const impersonate = createCheckbox('impersonate', isCreateMode ? 'Create as another user' : 'Update as another user');
        const rememberSettings = createCheckbox('remember', 'Remember these settings for this environment');

        const impersonationField = el('select', { on: { change: () => setImpersonationField(impersonationField.value) } },
            el('option', { value: 'systemuserid' }, 'systemuserid'),
            el('option', { value: 'azureactivedirectoryobjectid' }, 'azureactivedirectoryobjectid'));
        const impersonationLabel = el('span', { className: 'impersonationIdFieldLabel' }, 'systemuserid:');
        const impersonationInput = el('input', { placeholder: emptyGuid, on: { input: () => updateImpersonationPreview() } });
        const impersonationPreview = el('span', { className: 'impersonateUserPreview' });

        const impersonateDiv = el('div', { className: 'impersonateDiv pfwa-hidden' },
            el('div', {}, 'Base impersonation on this field: ', impersonationField, '  ', el('i', {}, el('a', { href: impersonationDocsUrl, target: '_blank' }, "What's this?"))),
            el('div', {}, impersonationLabel, impersonationInput, '  ', impersonationPreview));

        const editLink = actionLink('edit', () => editRecord(logicalName, pluralName, id, isCreateMode));
        editLink.title = 'Edit the columns of this row';

        const saveLink = actionLink('save', async () => {
            setHeaderState('saving');
            await submitEdit(pluralName, id, isCreateMode);
        }, 'primaryLink');
        const cancelLink = actionLink('cancel', () => reloadPage(pluralName));
        // below the changes, away from 'save' in the header, so a double click on 'save' does not commit
        const commitLink = actionLink('commit save', async () => {
            const { changes, impersonateHeader } = menu.pendingSave;
            await commitSave(pluralName, id, changes, impersonateHeader, isCreateMode, true);
        }, 'primaryLink');
        commitLink.title = 'Save the changes listed above';
        const backLink = actionLink('back', () => resetSubmitControls());
        backLink.title = 'Go back to editing';

        const changeCount = createMuted('');

        const savingNote = createMuted(' · saving…');

        // The header always shows 'save · cancel' while editing, so nothing moves under the cursor:
        // editing: both active, confirm: save disabled (committing is below the changes),
        // saving: both disabled, until the page is replaced (after a successful save) or the save fails and editing resumes.
        function setHeaderState(headerState) {
            const isSaving = headerState === 'saving';
            const isCommitting = isSaving && menu.pendingSave != null;

            for (const link of [cancelLink, commitLink, backLink]) {
                link.classList.toggle('disabled', isSaving);
            }

            saveLink.classList.toggle('disabled', headerState !== 'editing');

            setVisible(savingNote, isSaving);
            setVisible(commitActions, headerState === 'confirm' || isCommitting);
        }

        setActions(header.editActions, [saveLink, cancelLink]);

        header.editContext.append((isCreateMode ? 'new ' : 'editing ') + logicalName, ' ', changeCount, savingNote);

        const previewContainer = el('div', { className: 'previewChangesDiv' });
        const commitActions = el('div', { className: 'commitActions actions pfwa-hidden' }, commitLink, backLink); // 'commit save · back', in the same order as 'save · cancel'

        const options = el('div', { className: 'editOptions' },
            [previewChanges, bypassPlugins, bypassFlows, impersonate, rememberSettings].map(checkbox => checkbox.label));

        header.editBody.append(options, impersonateDiv, previewContainer, commitActions);
        setHeaderState('editing');

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

        const menu = {
            editLink,
            changeCount,
            previewContainer,
            setHeaderState,
            pendingSave: null, // the changes which are previewed, saved when committing
            bypassPlugins: bypassPlugins.input,
            bypassFlows: bypassFlows.input,
            previewChanges: previewChanges.input,
            impersonate: impersonate.input,
            impersonationField,
            impersonationInput,
            rememberSettings: rememberSettings.input,
            setImpersonatedUser,
        };

        view.editMenu = menu;

        return menu;
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

        const editLink = actionLink('edit', () => beginEdit('update'));
        editLink.title = 'Point this lookup to another row';
        const clearLink = actionLink('clear', () => beginEdit('makenull'));
        clearLink.title = 'Clear this lookup (set it to null)';
        const undoLink = actionLink('undo', () => undo());
        undoLink.title = 'Undo the changes to this lookup';

        // e.g. 'edit · clear · undo', undo is only shown when the lookup was changed
        setVisible(undoLink, false);

        const links = el('span', { className: 'lookupEditLinks actions' }, editLink, clearLink, undoLink);

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
            setVisible(undoLink, true);
        }

        function undo() {
            editor.editMode = null;
            input.value = '';

            setVisible(undoLink, false);
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

        try {
            if (!isCreateMode) {
                await addMissingStateOrStatus(view, pluralName, id);
            }

            await createEditors(view, logicalName, isCreateMode);
        } catch (error) {
            // show the plain error, and cancel the edit once it is dismissed
            console.error(error);
            alert(error?.message ?? String(error));
            await reloadPage(pluralName);
            return;
        }

        view.panel.classList.add('editing');

        applySettings(view.editMenu);

        // keep the number of changes in the header up to date
        let countTimeout = null;
        const scheduleCount = () => {
            clearTimeout(countTimeout);
            countTimeout = setTimeout(() => updateChangeCount(view, isCreateMode), 150);
        };

        for (const eventName of ['input', 'change', 'click']) {
            view.panel.addEventListener(eventName, scheduleCount);
        }

        updateChangeCount(view, isCreateMode);
    }

    // A status belongs to a state, so statecode and statuscode are edited together: when the row shows only one of them,
    // the current value of the other is retrieved and shown next to it, only while editing.
    async function addMissingStateOrStatus(view, pluralName, id) {
        const record = state.originalRecord;
        const present = ['statecode', 'statuscode'].filter(name => name in record);

        if (present.length !== 1 || !view.slots.has(present[0])) {
            return;
        }

        const [presentName] = present;
        const missingName = presentName === 'statecode' ? 'statuscode' : 'statecode';

        const json = await odataFetch(`${apiUrl}${pluralName}(${id})?$select=${missingName}`, false);

        if (json.error) {
            throw new Error(json.error.message);
        }

        record[missingName] = json[missingName] ?? null; // the value the change is compared with

        const slotsBefore = [...view.slots];
        const row = createRow(createCopyableKey(missingName),
            await renderPropertyValue(json, missingName, { view }, new Set()),
            createMuted(`  not selected, edited together with ${presentName}`));
        row.classList.add('editOnly');

        // in the same place as when both are selected: statecode first
        const presentRow = view.slots.get(presentName).element.closest('.pfwa-row');
        const missingFirst = missingName === 'statecode';
        presentRow[missingFirst ? 'before' : 'after'](row);

        // the slots are in the displayed order, so that the changes are listed in that order
        const missingSlot = view.slots.get(missingName);
        view.slots = new Map(slotsBefore.flatMap(entry => entry[0] !== presentName
            ? [entry]
            : missingFirst ? [[missingName, missingSlot], entry] : [entry, [missingName, missingSlot]]));
    }

    // creates an editor for every editable column of the row
    async function createEditors(view, logicalName, isCreateMode) {
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
    }

    // e.g. '2 changes' or '1 change · invalid value'
    async function updateChangeCount(view, isCreateMode) {
        let count = 0;
        let hasInvalidValue = false;

        for (const editor of view.editors) {
            try {
                if (await readEditorChange(editor, isCreateMode)) {
                    count++;
                }
            } catch (error) {
                if (!(error instanceof UserError)) {
                    throw error;
                }
                hasInvalidValue = true;
            }
        }

        view.editMenu.changeCount.textContent = `${count} ${count === 1 ? 'change' : 'changes'}${hasInvalidValue ? ' · invalid value' : ''}`;
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

        menu.setHeaderState('editing');
        menu.pendingSave = null;
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

    // returns { column, payloadKey, oldValue, newValue } when the value of the editor changed, null otherwise
    // throws a UserError when the value is invalid
    async function readEditorChange(editor, isCreateMode) {
        if (editor.dataType === 'lookup') {
            const change = await editor.readChange();

            if (change == null || !hasChanged(change.value, change.originalValue, 'lookup')) {
                return null;
            }

            return {
                column: editor.navigationProperty,
                payloadKey: `${editor.navigationProperty}@odata.bind`,
                oldValue: change.originalValue,
                newValue: change.value
            };
        }

        // if not found for some reason, normalize to null to prevent issues
        const originalValue = isCreateMode ? null : (state.originalRecord[editor.fieldName] ?? null);
        const value = editor.read();

        return hasChanged(value, originalValue, editor.dataType)
            ? { column: editor.fieldName, payloadKey: editor.fieldName, oldValue: originalValue, newValue: value }
            : null;
    }

    // returns a list of { column, payloadKey, oldValue, newValue }
    async function collectChanges(editors, isCreateMode) {
        const changes = [];

        for (const editor of editors) {
            const change = await readEditorChange(editor, isCreateMode);

            if (change) {
                changes.push(change);
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

        // the changes are listed in the header, with 'commit save' below them
        view.editMenu.pendingSave = { changes, impersonateHeader };
        view.editMenu.setHeaderState('confirm');

        view.editMenu.previewContainer.replaceChildren(changes.length > 0
            ? createChangesTable(changes, isCreateMode)
            : el('div', {}, createMuted('no changes')));
    }

    function createChangesTable(changes, isCreateMode) {
        const table = el('table', { id: 'previewTable' });

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

        menu.setHeaderState('saving');

        let response;

        try {
            response = await fetch(isCreateMode ? apiUrl + pluralName : apiUrl + pluralName + '(' + id + ')', {
                method: isCreateMode ? 'POST' : 'PATCH',
                headers,
                body: JSON.stringify(Object.fromEntries(changes.map(change => [change.payloadKey, change.newValue])))
            });
        } catch (error) {
            // no response at all, e.g. offline. Nothing was saved, so editing resumes
            alert(error?.message ?? String(error));
            resetSubmitControls();
            return;
        }

        // the save failed, editing resumes with the changes kept
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
            await reloadPage(pluralName);
        }

        // the actions stay disabled until the page is replaced
    }

    async function handleCreateNewRecord(logicalName) {
        let creatableAttributes;

        try {
            creatableAttributes = await retrieveAttributes(logicalName, 'IsValidForCreate');
        } catch (error) {
            alert(error.message);
            return;
        }

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

    async function deleteRecord(pluralName, id) {
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

    // While previewing, the main panel scrolls by itself instead of the page. These keep the same rows in view when switching.
    function enterPreviewLayout(mainPanel) {
        if (mainPanel.classList.contains('prePreviewed')) {
            return; // already previewing another row, the main panel keeps its own scroll position
        }

        const scrolledPastPanelTop = -mainPanel.getBoundingClientRect().top;

        mainPanel.classList.add('prePreviewed');
        document.body.classList.add('bodyPreviewed');

        mainPanel.scrollTop = scrolledPastPanelTop;
    }

    function leavePreviewLayout(mainPanel) {
        const scrolledPastPanelTop = mainPanel.scrollTop;

        mainPanel.classList.remove('prePreviewed');
        document.body.classList.remove('bodyPreviewed');

        window.scrollTo(window.scrollX, mainPanel.getBoundingClientRect().top + window.scrollY + scrolledPastPanelTop);
    }

    async function previewRecord(pluralName, id) {
        const mainPanel = state.mainView.panel;
        enterPreviewLayout(mainPanel);

        const previewPanel = el('div', { className: 'panel previewPanel prePreviewed' });
        document.body.appendChild(previewPanel);

        const response = await odataFetch(apiUrl + pluralName + '(' + id + ')', false);

        const { header } = await renderPanel(response, previewPanel, pluralName, { isPreview: true });

        const closeLink = actionLink('close', () => {
            if (document.getElementsByClassName('previewPanel').length === 1) {
                leavePreviewLayout(mainPanel);
            }

            previewPanel.remove();
        });
        closeLink.title = 'Close this preview';

        setActions(header.readActions, [closeLink]);

        // only scroll sideways to bring the preview into view (when there are multiple), never the main panel
        previewPanel.scrollIntoView({ block: 'nearest', inline: 'nearest' });
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

        filterInput.focus();
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

    // Adds 'edit flow' to the actions in the header. It is disabled until the editor has loaded:
    // the editor runs in the page context (initMonaco.js), which activates the link and handles its clicks.
    function createMonacoEditorControls(mainPanel, header, recordId, clientdata) {
        if (checkIfJsonViewerEnabled()) {
            header.readActions.append(el('span', { className: 'muted', title: 'Currently, editing Flows is only possible in Chrome. Work in progress.' }, 'edit flow'));
            return;
        }

        header.readActions.append(el('a', {
            className: 'editFlowLink disabled',
            href: '#',
            title: 'Loading the flow editor…',
            on: { click: (e) => e.preventDefault() }
        }, 'edit flow'));

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

            /* the header: a line which sticks to the top of the panel, what is shown on the left, the actions on the page on the right */
            .pfwa-toolbar {
                position: sticky;
                top: 0;
                z-index: 10;
                display: flex;
                flex-wrap: wrap; /* in a narrow panel the actions move below what is shown, instead of out of view */
                align-items: baseline;
                gap: 2px 3ch;
                margin: -10px -10px 6px;
                padding: 8px 10px 6px;
                background: white;
                border-bottom: 1px solid rgba(128, 128, 128, 0.4);
            }

            .headerActions {
                margin-left: auto;
            }

            /* what is shown on the left takes the space next to the actions, and a long text is cut off with '…'.
               Only when even 20 characters do not fit, the actions move below it */
            .headerContext {
                flex: 1 1 20ch;
                min-width: 0;
                overflow: hidden;
                white-space: nowrap;
                text-overflow: ellipsis;
            }

            /* plain inline text, so that the '…' cuts off the name itself instead of hiding it as a whole */
            .headerContext .field, .headerContext .field > * {
                display: inline;
            }

            /* while editing, below the line: the options, the impersonation inputs and the changes to commit */
            .headerBody {
                flex-basis: 100%;
            }

            .headerBody > * {
                margin-top: 4px;
            }

            /* a long list of changes scrolls, so the header does not cover the whole panel */
            .previewChangesDiv {
                max-height: 40vh;
                overflow: auto;
            }

            .previewChangesDiv:empty {
                display: none;
            }

            @media (prefers-color-scheme: dark) {
                .pfwa-toolbar {
                    background: #28282B;
                }
            }

            /* the read parts are shown normally, the edit parts while editing */
            .editing .readOnly, .pfwa-root:not(.editing) .editOnly {
                display: none !important;
            }

            /* actions are separated with ' · ', drawn here so that a hidden action takes its separator with it */
            .actions > :not(.pfwa-hidden) ~ :not(.pfwa-hidden) {
                position: relative;
                margin-left: 3ch;
            }

            /* drawn in the gap in front of the action, outside of it, so it is not underlined and clicking it does not click the action */
            .actions > :not(.pfwa-hidden) ~ :not(.pfwa-hidden)::before {
                content: '·';
                position: absolute;
                left: -3ch;
                width: 3ch;
                text-align: center;
                color: gray;
                pointer-events: none;
            }

            /* the same dot between other parts, e.g. the table name and the primary name in the header */
            .separator {
                display: inline-block;
                width: 3ch;
                text-align: center;
                color: gray;
            }

            @media (prefers-color-scheme: dark) {
                .actions > :not(.pfwa-hidden) ~ :not(.pfwa-hidden)::before, .separator {
                    color: #8c8c8c;
                }
            }

            /* actions after a value on the same line */
            .inlineActions, .lookupEditLinks {
                margin-left: 2ch;
            }

            /* secondary text, e.g. the row count in the header */
            .muted {
                color: gray;
            }

            a.dangerLink {
                color: firebrick;
            }

            /* the actions which save, bold so they draw the eye (and in dark mode in the yellow of the logo) */
            a.primaryLink {
                font-weight: bold;
            }

            /* every active link or action changes colour on hover. In light mode to a deep, saturated blue,
               and delete to pure red. In dark mode it gets lighter, see below */
            a[href]:not(.busy, .disabled):hover {
                color: #0000aa;
            }

            a.dangerLink[href]:not(.busy, .disabled):hover {
                color: #ff0000;
            }

            button:not(:disabled):hover {
                filter: brightness(0.92);
            }

            /* links which open a new tab get the familiar icon, drawn in the colour of the link */
            a[target="_blank"][href]::after {
                content: '';
                display: inline-block; /* not underlined with the link */
                width: 0.8em;
                height: 0.8em;
                margin-left: 0.25em;
                vertical-align: -0.05em;
                background-color: currentColor;
                -webkit-mask: url("data:image/svg+xml,%3Csvg xmlns='http://www.w3.org/2000/svg' viewBox='0 0 24 24'%3E%3Cpath d='M19 19H5V5h7V3H5a2 2 0 0 0-2 2v14a2 2 0 0 0 2 2h14a2 2 0 0 0 2-2v-7h-2v7zM14 3v2h3.59l-9.83 9.83 1.41 1.41L19 6.41V10h2V3h-7z'/%3E%3C/svg%3E") no-repeat center / contain;
                mask: url("data:image/svg+xml,%3Csvg xmlns='http://www.w3.org/2000/svg' viewBox='0 0 24 24'%3E%3Cpath d='M19 19H5V5h7V3H5a2 2 0 0 0-2 2v14a2 2 0 0 0 2 2h14a2 2 0 0 0 2-2v-7h-2v7zM14 3v2h3.59l-9.83 9.83 1.41 1.41L19 6.41V10h2V3h-7z'/%3E%3C/svg%3E") no-repeat center / contain;
            }

            /* an action which is running or disabled, see actionLink */
            a.busy, a.disabled, a.dangerLink.busy, a.dangerLink.disabled, a.primaryLink.busy, a.primaryLink.disabled {
                color: gray;
                cursor: progress;
            }

            @media (prefers-color-scheme: dark) {
                .muted {
                    color: #8c8c8c;
                }

                a.dangerLink {
                    color: #ff5c5c;
                }

                /* the yellow of the logo, softened */
                a.primaryLink {
                    color: #e3dc6a;
                }

                a.busy, a.disabled, a.dangerLink.busy, a.dangerLink.disabled, a.primaryLink.busy, a.primaryLink.disabled {
                    color: #8c8c8c;
                }

                a[href]:not(.busy, .disabled):hover {
                    color: #8ab4f8;
                }

                a.dangerLink[href]:not(.busy, .disabled):hover {
                    color: #ff7575;
                }

                /* the same yellow, paler */
                a.primaryLink[href]:not(.busy, .disabled):hover {
                    color: #efeaa0;
                }

                button:not(:disabled):hover {
                    filter: brightness(1.6);
                }
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
            .primarykey { color: #0000EE; }

            @media (prefers-color-scheme: dark) {
            *:not(svg, path, a) {
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

            .panel .field:not(.lookupField, .inlineField), .panel .lookupDisplay {
                margin-right: 24px;
                padding-right: 16px;
            }

            /* the formatted value is secondary to the raw value, so it is dimmed */
            .field .formattedValue {
                color: gray;
            }

            /* the part of a field which is copied, it keeps the color of the field (also in dark mode) */
            .field .copyValue {
                color: inherit;
            }

            /* the whole field is clickable, but only the part which is copied is underlined on hover and flashes when copied */
            .panel:not(.editing) .field:hover .copyValue, .panel:not(.editing) .copyableKey:hover {
                text-decoration: underline dotted;
            }

            .panel:not(.editing) .copyableKey {
                cursor: pointer;
            }

            .field.copied .copyValue, .copyableKey.copied {
                background: rgba(128, 128, 128, 0.35);
            }

            /* a very long value is cut off at the edge of the page with '…', followed by 'show all', see createField.
               Its row becomes a flex row, so the value can shrink to the space left after the column name */
            .pfwa-row:has(> .longValue:not(.showingAll)), .pfwa-row > b:has(> .longValue:not(.showingAll)) {
                display: flex;
            }

            .longValue:not(.showingAll), .pfwa-row > b:has(> .longValue:not(.showingAll)) {
                min-width: 0;
            }

            .longValue:not(.showingAll) > .copyValue {
                display: block; /* also cut off while editing, when a memo field is not a flex box */
                min-width: 0;
                overflow: hidden;
                white-space: nowrap;
                text-overflow: ellipsis;
            }

            .longValue > :not(.copyValue) {
                flex: none;
            }

            .showAllLink {
                margin-left: 1ch;
            }

            /* while editing, the editor shows all of the value */
            .editing .showAllLink {
                display: none;
            }

            /* While editing a long memo value, its textarea goes on a line of its own below the row, at the same distance
               from the left as the textarea of a short memo value. The parts of the field become parts of the flex row,
               with an empty line break (the ::after of the row) between the value and the textarea */
            .editing .pfwa-row:has(> .longValue.memoField:not(.showingAll)) {
                flex-wrap: wrap;
                padding-bottom: 1lh; /* the same space below the textarea as a short memo value has (an empty line) */
            }

            .editing .pfwa-row:has(> .longValue.memoField:not(.showingAll))::after {
                content: '';
                flex-basis: 100%;
                order: 1;
            }

            .editing .longValue.memoField:not(.showingAll) {
                display: contents;
            }

            .editing .longValue.memoField:not(.showingAll) > .copyValue {
                flex: 1 1 0;
            }

            .editing .longValue.memoField:not(.showingAll) > textarea {
                order: 2;
            }

            /* all of a long value, wrapped to the width of the page instead of one very long line */
            .field.showingAll {
                display: inline;
            }

            .field.showingAll .copyValue {
                white-space: pre-wrap;
                overflow-wrap: anywhere;
            }

            @media (prefers-color-scheme: dark) {
                .field .formattedValue {
                    color: #8c8c8c;
                }
            }

            /* edit mode */

            .lookupEdit {
                display: none;
            }

            .editing .lookupEdit {
                display: unset;
            }

            .editing .lookupDisplay, .locked .lookupEditLinks {
                display: none;
            }

            .editing .field {
                cursor: auto;
                height: 20px;
            }

            /* the textarea goes below the value, and the row grows with it.
               (A long value makes its row a flex row, in which the field is a block, so it needs the height as well) */
            .editing .field.memoField {
                display: inline;
                height: auto;
            }

            .editing .field.memoField textarea {
                display: block;
            }

            .editing .lookupField {
                margin: 1px 0 -2px 0;
            }

            /* the edit options, the first line of the main row while editing, e.g. '[x] preview changes  [ ] bypass plugins' */
            .editOption {
                margin-right: 2ch;
                cursor: pointer;
            }

            .panel .editOption input {
                cursor: pointer; /* checkboxes don't inherit the cursor of their label */
                width: auto;
                margin: 0 4px 0 0;
                vertical-align: -2px;
            }

            /* while the changes are previewed, the options are disabled */
            .locked .editOption, .panel .editOption input:disabled {
                cursor: default;
            }

            .lookupEditRow, .makeLookupNullDiv {
                padding-left: 3ch;
            }

            .lookupSelectTable {
                height: 22px;
            }

            .lookupPreview {
                margin-left: 10px;
            }

            /* changes preview table, in the header */

            .panel table {
                color: black;
                margin: 0;
                border-collapse: collapse;
                border: 1px solid black;
                table-layout: fixed;
                width: 100%;
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

            .bodyPreviewed {
                display: inline-flex;
                margin-top: 0px;
                margin-bottom: 0px;
            }

            /* side by side panels which scroll independently, so their toolbars stay in view */
            .prePreviewed {
                width: 48vw;
                overflow-x: scroll !important;
                overflow-y: scroll !important;
                height: 100vh;
                box-sizing: border-box;
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

    document.addEventListener('mouseover', showCopyTooltip);

    await makeItPretty();
})()
