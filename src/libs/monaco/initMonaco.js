(function () {
    const body = document.body;
    const mainPanel = document.getElementsByClassName('mainPanel')[0];
    const chromeRuntimeUrl = mainPanel.dataset.chromeRuntimeUrl;

    const container = document.createElement('section');
    container.style.display = 'none';
    container.classList.add("monacoContainer");
    body.appendChild(container);

    container.innerHTML = `
<div id="regularEditorContainer" class="monacoContainer"></div>
<div id="diffEditorContainer" class="monacoContainer" style="display: none;"></div>
                
<div class="monacoActions">
  <label class="monacoLabel" id="monacoLabel" style="cursor:pointer;">Show differences</label>
  <input type="checkbox" style="width:unset;cursor:pointer;" class="showDifferenceInput" id="showDifferenceInput">&nbsp;&nbsp;
  <button id="prettifyJsonButton" style="cursor:pointer;">Prettify JSON</button>&nbsp;&nbsp;
  <button id="renameThingsButton" style="cursor:pointer;">Rename variable/action</button>&nbsp;&nbsp;
  <button id="renameTextButton" style="cursor:pointer;">Find & Replace</button>&nbsp;&nbsp;
  <button id="backupFlowButton" style="cursor:pointer;">Create backup</button>&nbsp;&nbsp;
  <button id="saveFlowButton" style="cursor:pointer;">Save Flow</button>&nbsp;&nbsp;
  <button id="cancelEditFlowButton" style="cursor:pointer;">Close editor</button>
</div>`;

    const btn = document.createElement('button');
    btn.style = `
        height: 30px;
        width: auto;
        margin-right: 160px;
        margin-top: 10px;
        position: absolute;
        right: 10px;
        cursor: pointer;
        padding:0;
        font-size:24;
        padding: 0px 4px;
        display:none;
        `
    btn.innerHTML = '<div>Edit Flow</div>';

    mainPanel.prepend(btn);

    let regularEditor = null;
    let diffEditor = null;

    function prettifyJsonString(json) {
        return JSON.stringify(JSON.parse(json), null, 2);
    }

    function createRegularEditor(value) {
        regularEditor = monaco.editor.create(document.getElementById("regularEditorContainer"), {
            value: value,
            language: "json",
            automaticLayout: true,
        });

        window.regularEditor = regularEditor;
    }

    function setRegularEditorValue(value) {
        regularEditor.getModel().setValue(value);
    }

    function createDiffEditor(original) {
        diffEditor = monaco.editor.createDiffEditor(document.getElementById('diffEditorContainer'), {
            automaticLayout: true,
            renderSideBySide: true
        });
        diffEditor.setModel({
            original: monaco.editor.createModel(original, 'json'),
            modified: monaco.editor.createModel(original, 'json'),
        });

        window.diffEditor = diffEditor;
    }

    function setDiffEditorValue(modified) {
        diffEditor.getModel().modified.setValue(modified);
    }

    function setDiffEditorOriginalValue(original) {
        diffEditor.getModel().original.setValue(original);
    }

    function backupFlow(name, text) {
        const element = document.createElement('a');
        element.setAttribute('href', 'data:text/plain;charset=utf-8,' + encodeURIComponent(text));
        element.setAttribute('download', name);

        element.style.display = 'none';
        document.body.appendChild(element);

        element.click();

        document.body.removeChild(element);
    }

    function getEditorValue() {
        if (editMode === 'regular') {
            return regularEditor.getModel().getValue();
        } else if (editMode === 'diff') {
            return diffEditor.getModel().modified.getValue();
        } else {
            const err = 'Invalid editmode. Should not happen.';
            alert(err);
            throw err;
        }
    }

    async function commitSave(text) {
        const requestUrl = mainPanel.dataset.apiUrl + 'workflows(' + mainPanel.dataset.recordId + ')';

        let headers = {
            'accept': 'application/json',
            'content-type': 'application/json',
            'OData-MaxVersion': '4.0',
            'OData-Version': '4.0',
            'If-Match': '*'
        };

        const response = await fetch(requestUrl, {
            method: 'PATCH',
            headers: headers,
            body: JSON.stringify({ clientdata: text })
        });

        if (response.ok) {
            alert('Changes saved!');
            originalValue = text;
            setDiffEditorOriginalValue(text);
        } else {
            const errorText = await response.text();
            console.error(`${response.status} - ${errorText}`);
            window.alert(`${response.status} - ${errorText}`);
        }
    }

    let editMode = 'regular';
    let originalValue = JSON.stringify(JSON.parse(mainPanel.dataset.clientdata), null, 2);

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

        document.getElementById('monacoLabel').onclick = async () => {
            document.getElementById('showDifferenceInput').click();
        }

        document.querySelector('.showDifferenceInput').addEventListener('change', (e) => {
            if (e.target.checked) {
                editMode = 'diff';
                setDiffEditorValue(regularEditor.getModel().getValue());
                document.getElementById('regularEditorContainer').style.display = 'none'
                document.getElementById('diffEditorContainer').style.display = null
            } else {
                editMode = 'regular';
                setRegularEditorValue(diffEditor.getModel().modified.getValue());
                document.getElementById('regularEditorContainer').style.display = null
                document.getElementById('diffEditorContainer').style.display = 'none'
            }
        });

        document.getElementById("prettifyJsonButton").onclick = () => {
            if (editMode === 'regular') {
                regularEditor.getAction("editor.action.formatDocument").run();;
            } else if (editMode === 'diff') {
                diffEditor.getModifiedEditor().getAction('editor.action.formatDocument').run()
            } else {
                const err = 'Invalid editmode. Should not happen.';
                alert(err);
                throw err;
            }
        };

        btn.style.display = null;
        btn.onclick = async () => {
            container.style.display = null;

            const panels = document.getElementsByClassName('panel');
            for (let i = 0; i < panels.length; i++) {
                panels[i].style.display = 'none';
            }
        };

        document.getElementById('saveFlowButton').onclick = async () => {
            if (confirm('Save this flow?')) {
                await commitSave(getEditorValue());
            }
        }

        document.getElementById('backupFlowButton').onclick = () => {
            let name = mainPanel.dataset.flowName;

            if (name == null || name == 'undefined') {
                name = mainPanel.dataset.recordId;
            }

            backupFlow(name + '.json', originalValue);
        }

        document.getElementById('cancelEditFlowButton').onclick = async () => {
            window.location.reload();
        }

        document.getElementById('renameThingsButton').onclick = async () => {
            let selection = '';
            if (editMode === 'regular') {
                selection = regularEditor.getModel().getValueInRange(regularEditor.getSelection())
            } else if (editMode === 'diff') {
                selection = diffEditor.getModifiedEditor().getModel().getValueInRange(diffEditor.getModifiedEditor().getSelection());
            } else {
                const err = 'Invalid editmode. Should not happen.';
                alert(err);
                throw err;
            }

            if (selection == null || selection === '') {
                alert('To use this method, first select the action or variable name to replace in the editor.');
                return;
            }

            selection = selection.replaceAll("'", '').replaceAll('"', '');

            const newName = prompt('Please enter a new name for : ' + selection, selection);
            if (newName != null && newName !== '') {
                const currentValue = getEditorValue();
                const newValue = currentValue.replaceAll(`'${selection}'`, `'${newName}'`).replaceAll(`"${selection}"`, `"${newName}"`);

                if (editMode === 'regular') {
                    setRegularEditorValue(newValue);
                } else if (editMode === 'diff') {
                    setDiffEditorValue(newValue);
                } else {
                    const err = 'Invalid editmode. Should not happen.';
                    alert(err);
                    throw err;
                }
            }
        }

        document.getElementById('renameTextButton').onclick = async () => {
            if (editMode === 'regular') {
                regularEditor.getAction('editor.action.startFindReplaceAction').run();;
            } else if (editMode === 'diff') {
                diffEditor.getModifiedEditor().getAction('editor.action.startFindReplaceAction').run()
            } else {
                const err = 'Invalid editmode. Should not happen.';
                alert(err);
                throw err;
            }
        }
    });
})()