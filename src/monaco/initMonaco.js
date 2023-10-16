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
  <label class="monacoLabel">Show differences</label>
  <input type="checkbox" style="width:unset;" class="showDifferenceInput" >&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
  <button id="prettifyJsonButton">Prettify JSON</button>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
  <button id="backupFlowButton">Create backup</button>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
  <button id="saveFlowButton">Save Flow</button>
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

    document.getElementById("prettifyJsonButton").onclick = () => {
        setRegularEditorValue(prettifyJsonString(regularEditor.getValue()));
        setDiffEditorValue(prettifyJsonString(diffEditor.getModel().modified.getValue()));
    };

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
    }

    function setDiffEditorValue(modified) {
        diffEditor.getModel().modified.setValue(modified);
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
        } else {
            const errorText = await response.text();
            console.error(`${response.status} - ${errorText}`);
            window.alert(`${response.status} - ${errorText}`);
        }
    }

    let editMode = 'regular';

    window.setTimeout(() => {
        require.config({ paths: { vs: chromeRuntimeUrl + 'monaco' } });

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

            const originalValue = JSON.stringify(JSON.parse(mainPanel.dataset.clientdata), null, 2);

            createRegularEditor(originalValue);
            createDiffEditor(originalValue);

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

                backupFlow(name + '.json', getEditorValue());
            }
        });
    }, 2000);
})()