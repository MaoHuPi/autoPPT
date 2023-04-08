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
        if(pageNow < pageMax && pageNow < pages.childElementCount-1) floatButtonNext.removeAttribute('disabled');
        else floatButtonNext.setAttribute('disabled', '');
        if(pageNow > 0) floatButtonBack.removeAttribute('disabled');
        else floatButtonBack.setAttribute('disabled', '');
    }
    updateFloatButton();
    floatButtonBack.addEventListener('click',  pageChange.bind({rate: -1}));
    floatButtonNext.addEventListener('click',  pageChange.bind({rate: 1}));

    // page-upload
    let settings = {};
    const docxButton = $('#docxButton');
    const docxViewer = $('#docxViewer');
    async function uploadAndAnalyzeDocx(){
        await ipcRenderer.invoke('uploadDocx')
        .then(data => {
            if(data.html !== undefined){
                docxViewer.innerHTML = data.html;
                docxViewer.removeAttribute('empty');
                pageMax = 1;
                updateFloatButton();
            }
            if(data.settings !== undefined){
                settings = data.settings;
                createSettingsForm(data.settings);
                pageMax = 2;
            }
            return(data);
        });
    }
    docxButton.addEventListener('click', uploadAndAnalyzeDocx);

    // page-choose
    const pptxButton = $('#pptxButton');
    async function exportPptx(){
        await ipcRenderer.invoke('exportPptx', settings);
    }
    pptxButton.addEventListener('click', exportPptx);

    const settingsForm = $('#settingsForm');
    function createSettingsForm(settindsFormData){
        settingsForm.innerHTML = '';
        settingsForm.removeAttribute('empty');
        let blockNum = 2;
        let inputPerBlock = Math.round(Object.keys(settindsFormData).length / blockNum);
        let blockCount = 0;
        let blockNow = false;
        for(let name in settindsFormData){
            if(!(blockNow) || (blockNow && blockNow.childElementCount >= inputPerBlock)){
                blockNow = $e('div');
                blockNow.className = 'block';
                blockNow.style.setProperty('--blockNum', blockNum);
                blockNow.style.setProperty('--hasMarginLeft', ++blockCount % blockNum == 0 ? 0 : 1);
                settingsForm.appendChild(blockNow);
            }
            let type = settindsFormData[name].type.split('(');
            let value = settindsFormData[name].value;
            let inputAttribute = type.length > 1 ? type[1].replace(/\)$/, '').replaceAll(', ', ',').split(',') : [];
            type = type[0].toLowerCase();;
            let inputId = `settings-${name}`;
            let row = $e('div');
            let label = $e('label');
            let input = $e(type.indexOf('select') > -1 ? 'select' : 'input');
            row.className = 'inputRow';
            label.innerText = name;
            label.setAttribute('for', inputId);
            input.id = inputId;
            input.name = inputId;
            let regulateFunction = value => value;
            switch(type){
                case 'int':
                    regulateFunction = value => parseInt(parseFloat(value));
                case 'float':
                    regulateFunction = value => parseFloat(value);
                case 'num':
                case 'number':
                    inputAttribute = inputAttribute.map(attr => Number(attr));
                    value = Number(value);
                    ['min', 'max', 'step'].map((attrName, i) => {
                        if(inputAttribute[i] !== undefined) input.setAttribute(attrName, inputAttribute[i]);
                    });
                    let oldRegulateFunction = regulateFunction;
                    regulateFunction = v => {
                        v = oldRegulateFunction(v);
                        v = Number(v);
                        if(inputAttribute[0] !== undefined && v < inputAttribute[0]) v = inputAttribute[0];
                        if(inputAttribute[1] !== undefined && v > inputAttribute[1]) v = inputAttribute[1];
                        if(inputAttribute[2] !== undefined){
                            var startValue = value !== undefined ? value : inputAttribute[0] !== undefined ? inputAttribute[0] : 0;
                            var valueRate = inputAttribute[2];
                            v = Math.round((v-startValue)/valueRate) * valueRate + startValue;
                        }
                        return(v);
                    };
                    input.type = 'number';
                    input.value = value;
                    break;
                case 'date':
                    if(inputAttribute[0] !== undefined) input.setAttribute('min', inputAttribute[0]);
                    if(inputAttribute[1] !== undefined) input.setAttribute('max', inputAttribute[1]);
                    if(inputAttribute[2] !== undefined) input.setAttribute('step', inputAttribute[2]);
                    input.type = 'date';
                    input.value = value;
                    break;
                case 'check':
                case 'checkbox':
                    input.type = 'checkbox';
                    value = [true, 1, 'true'].indexOf(value) > -1 ? true : false;
                    input.checked = value;
                    settings[name].value = value;
                    break;
                case 'selector':
                case 'select':
                    inputAttribute.forEach(attr => {
                        let option = $e('option');
                        option.innerText = attr;
                        option.value = attr;
                        input.appendChild(option);
                    });
                    input.value = value;
                    break;
                default:
                    input.type = type;
                    input.value = value;
                    break;
            }
            function onChange(){
                this.value = regulateFunction(this.value);
                settings[name].value = this.tagName.toLowerCase() == 'input' && this.getAttribute('type').toLowerCase() == 'checkbox' ? 
                this.checked : this.value;
            }
            input.addEventListener('blur', onChange);
            input.addEventListener('change', onChange);
            row.appendChild(label);
            row.appendChild(input);
            blockNow.appendChild(row);
        }
    }
});