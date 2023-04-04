const { contextBridge, ipcRenderer} = require('electron');

window.addEventListener('DOMContentLoaded', () => {
    window.info = {};
    for (let type of ['chrome', 'node', 'electron']) {
        window.info[`${type}-version`] = process.versions[type];
    }

    let docxButton = document.querySelector('#docxButton');
    // contextBridge.exposeInMainWorld('electronAPI', {
    //     uploadDocx: () => ipcRenderer.send('uploadDocx')
    // })
    async function uploadAndAnalyzeDocx(){
        let $ = (e, p = document) => p.querySelector(e);
        let $$ = (e, p = document) => p.querySelectorAll(e);
        let $e = document.createElement.bind(document);
        await ipcRenderer.invoke('uploadDocx')
        .then(data => {
            return(data);
        });
    }
    docxButton.addEventListener('click', uploadAndAnalyzeDocx);
});