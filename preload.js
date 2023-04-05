const { contextBridge, ipcRenderer} = require('electron');

window.addEventListener('DOMContentLoaded', () => {
    // basic
    let $ = (e, p = document) => p.querySelector(e);
    let $$ = (e, p = document) => p.querySelectorAll(e);
    let $e = document.createElement.bind(document);
    contextBridge.exposeInMainWorld('versions', process.versions);

    // window
    const html = $('html');
    const pages = $('#pages');
    let pageNow = 0;
    let pageMax = 0;
    const floatButtonBack = $('#floatButton-back');
    const floatButtonNext = $('#floatButton-next');
    function pageChange(rate){
        if(typeof rate != typeof 0) rate = this.rate; 
        if(pageNow + rate >= 0 && pageNow + rate < pages.childElementCount) pageNow += rate;
        html.style.setProperty('--pageNow', pageNow);
        updateFloatButton();
    }
    function updateFloatButton(){
        console.log(pageMax);
        if(pageNow < pageMax && pageNow < pages.childElementCount-1) floatButtonNext.removeAttribute('disabled');
        else floatButtonNext.setAttribute('disabled', '');
        if(pageNow > 0) floatButtonBack.removeAttribute('disabled');
        else floatButtonBack.setAttribute('disabled', '');
    }
    updateFloatButton();
    floatButtonBack.addEventListener('click',  pageChange.bind({rate: -1}));
    floatButtonNext.addEventListener('click',  pageChange.bind({rate: 1}));

    // page-upload
    const docxButton = $('#docxButton');
    async function uploadAndAnalyzeDocx(){
        await ipcRenderer.invoke('uploadDocx')
        .then(html => {
            if(html !== undefined){
                $('#docxViewer').innerHTML = html;
                pageMax = 1;
                updateFloatButton();
            }
            return(html);
        });
    }
    docxButton.addEventListener('click', uploadAndAnalyzeDocx);

    // page-choose
    const pptxButton = $('#pptxButton');
    async function exportPptx(){
        await ipcRenderer.invoke('exportPptx');
    }
    pptxButton.addEventListener('click', exportPptx);
});