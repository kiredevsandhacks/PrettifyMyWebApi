(function () {
    const mainPanel = document.getElementsByClassName('mainPanel')[0];
    const { chromeRuntimeUrl } = mainPanel.dataset;

    // Creates an element. Children can be nodes or strings. Options: className, style, on (event handlers), and any other element property.
    function el(tag, options = {}, ...children) {
        const { className, style, on, ...properties } = options;
        const element = document.createElement(tag);

        if (className) {
            element.className = className;
        }
        if (style) {
            Object.assign(element.style, style);
        }
        for (const [eventName, handler] of Object.entries(on ?? {})) {
            element.addEventListener(eventName, handler);
        }

        element.append(...children);
        Object.assign(element, properties);

        return element;
    }

    function setVisible(element, visible) {
        element.style.display = visible ? '' : 'none';
    }

    const regularEditorContainer = el('div', { id: 'regularEditorContainer', className: 'monacoContainer' });
    const diffEditorContainer = el('div', { id: 'diffEditorContainer', className: 'monacoContainer', style: { display: 'none' } });

    const showDifferenceInput = el('input', { type: 'checkbox', id: 'showDifferenceInput', className: 'showDifferenceInput' });
    const prettifyJsonButton = el('button', { type: 'button' }, 'Prettify JSON');
    const renameThingsButton = el('button', { type: 'button' }, 'Rename variable/action');
    const renameTextButton = el('button', { type: 'button' }, 'Find & Replace');
    const backupFlowButton = el('button', { type: 'button' }, 'Create backup');
    const saveFlowButton = el('button', { type: 'button' }, 'Save Flow');
    const cancelEditFlowButton = el('button', { type: 'button' }, 'Close editor');

    const container = el('section', { className: 'monacoContainer', style: { display: 'none' } },
        regularEditorContainer,
        diffEditorContainer,
        el('div', { className: 'monacoActions' },
            el('label', { className: 'monacoLabel', htmlFor: 'showDifferenceInput' }, 'Show differences'),
            showDifferenceInput,
            prettifyJsonButton,
            renameThingsButton,
            renameTextButton,
            backupFlowButton,
            saveFlowButton,
            cancelEditFlowButton));

    document.body.appendChild(container);

    // the 'edit flow' action in the header, disabled until monaco has loaded
    const editFlowLink = document.querySelector('.editFlowLink');

    let regularEditor = null;
    let diffEditor = null;
    let editMode = 'regular';
    let originalValue = JSON.stringify(JSON.parse(mainPanel.dataset.clientdata), null, 2);

    function getTheme() {
        return window.matchMedia?.('(prefers-color-scheme: dark)').matches ? 'vs-dark' : 'vs';
    }

    function createRegularEditor(value) {
        regularEditor = monaco.editor.create(regularEditorContainer, {
            value: value,
            language: 'json',
            automaticLayout: true,
            theme: getTheme()
        });

        window.regularEditor = regularEditor;
    }

    function createDiffEditor(original) {
        diffEditor = monaco.editor.createDiffEditor(diffEditorContainer, {
            automaticLayout: true,
            renderSideBySide: true,
            theme: getTheme()
        });
        diffEditor.setModel({
            original: monaco.editor.createModel(original, 'json'),
            modified: monaco.editor.createModel(original, 'json'),
        });

        window.diffEditor = diffEditor;
    }

    // the editor the user is currently typing in
    function getActiveEditor() {
        if (editMode === 'regular') {
            return regularEditor;
        } else if (editMode === 'diff') {
            return diffEditor.getModifiedEditor();
        }

        const err = 'Invalid editmode. Should not happen.';
        alert(err);
        throw err;
    }

    function getEditorValue() {
        return getActiveEditor().getModel().getValue();
    }

    function setEditorValue(value) {
        getActiveEditor().getModel().setValue(value);
    }

    function runEditorAction(actionId) {
        getActiveEditor().getAction(actionId).run();
    }

    function downloadFile(name, text) {
        const link = el('a', {
            href: 'data:text/plain;charset=utf-8,' + encodeURIComponent(text),
            download: name,
            style: { display: 'none' }
        });

        document.body.appendChild(link);
        link.click();
        link.remove();
    }

    async function commitSave(text) {
        const originalButtonText = saveFlowButton.textContent;

        saveFlowButton.disabled = true;
        saveFlowButton.textContent = 'Saving...';

        // a link resolves the url the same way fetch does
        const request = el('a', { href: mainPanel.dataset.apiUrl + 'workflows(' + mainPanel.dataset.recordId + ')' });
        const requestUrl = request.href;

        // the flow is saved to the Web Api of this page (same origin: protocol, host and port, for on-prem)
        if (request.origin !== window.location.origin) {
            alert('Could not determine the Web Api url of this environment, the flow was not saved.');
            saveFlowButton.disabled = false;
            saveFlowButton.textContent = originalButtonText;
            return;
        }

        const response = await fetch(requestUrl, {
            method: 'PATCH',
            headers: {
                'accept': 'application/json',
                'content-type': 'application/json',
                'OData-MaxVersion': '4.0',
                'OData-Version': '4.0',
                'If-Match': '*'
            },
            body: JSON.stringify({ clientdata: text })
        });

        if (response.ok) {
            alert('Changes saved!');
            originalValue = text;
            diffEditor.getModel().original.setValue(text);
        } else {
            const errorText = await response.text();
            console.error(`${response.status} - ${errorText}`);
            window.alert(`${response.status} - ${errorText}`);
        }

        saveFlowButton.disabled = false;
        saveFlowButton.textContent = originalButtonText;
    }

    function toggleDiffMode() {
        // carry the current value over to the editor that is about to be shown
        const value = getEditorValue();

        editMode = showDifferenceInput.checked ? 'diff' : 'regular';
        setEditorValue(value);

        setVisible(regularEditorContainer, editMode === 'regular');
        setVisible(diffEditorContainer, editMode === 'diff');
    }

    function renameSelection() {
        const activeEditor = getActiveEditor();
        let selection = activeEditor.getModel().getValueInRange(activeEditor.getSelection());

        if (selection == null || selection === '') {
            alert('To use this method, first select the action or variable name to replace in the editor.');
            return;
        }

        selection = selection.replaceAll("'", '').replaceAll('"', '');

        const newName = prompt('Please enter a new name for : ' + selection, selection);
        if (newName != null && newName !== '') {
            const newValue = getEditorValue()
                .replaceAll(`'${selection}'`, `'${newName}'`)
                .replaceAll(`"${selection}"`, `"${newName}"`);

            setEditorValue(newValue);
        }
    }

    function openEditor() {
        setVisible(container, true);

        for (const panel of document.getElementsByClassName('panel')) {
            panel.style.display = 'none';
        }
    }

    function backupFlow() {
        let name = mainPanel.dataset.flowName;

        if (name == null || name == 'undefined') {
            name = mainPanel.dataset.recordId;
        }

        downloadFile(name + '.json', originalValue);
    }

    require.config({ paths: { vs: chromeRuntimeUrl + 'libs/monaco' } });

    require(['vs/editor/editor.main'], function () {
        monaco.languages.json.jsonDefaults.setDiagnosticsOptions({
            validate: true,
            schemas: [
                {
                    fileMatch: ['*'],
                    schema: {}
                }
            ]
        });

        createRegularEditor(originalValue);
        createDiffEditor(originalValue);

        showDifferenceInput.addEventListener('change', toggleDiffMode);
        prettifyJsonButton.addEventListener('click', () => runEditorAction('editor.action.formatDocument'));
        renameTextButton.addEventListener('click', () => runEditorAction('editor.action.startFindReplaceAction'));
        renameThingsButton.addEventListener('click', renameSelection);
        backupFlowButton.addEventListener('click', backupFlow);

        saveFlowButton.addEventListener('click', async () => {
            if (confirm('Save this flow?')) {
                await commitSave(getEditorValue());
            }
        });

        // monaco throws up when re-initializing so just reload
        // also, reloading makes sure we don't end up with invalid clientdata or anything
        cancelEditFlowButton.addEventListener('click', () => window.location.reload());

        editFlowLink.addEventListener('click', openEditor);
        editFlowLink.classList.remove('disabled');
        editFlowLink.title = 'Edit the definition of this flow';
    });
})()
