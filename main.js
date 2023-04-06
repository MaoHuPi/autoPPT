// basic import
const path = require('path');
const { app, BrowserWindow, Tray, ipcMain } = require('electron');

// module import
const fs = require('fs');
const { dialog } = require('electron');
const decompress = require('decompress');
var { parse } = require('node-html-parser');
const HTMLparser = parse;
const { XMLParser, XMLBuilder, XMLValidator} = require('fast-xml-parser');
const getImageSize = require('image-size');

// basic method
let $ = (e, p = document) => p.querySelector(e.toLocaleLowerCase());
let $$ = (e, p = document) => p.querySelectorAll(e.toLocaleLowerCase());
function deleteDir(dirPath){
    if(fs.existsSync(dirPath)){
        let filePathList = fs.readdirSync(dirPath) || [];
        filePathList.forEach(filePath => {
            filePath = path.join(dirPath, filePath); 
            if(fs.statSync(filePath).isDirectory()) deleteDir(filePath);                
            else{
                try{
                    fs.unlinkSync(filePath);
                }
                catch(err){console.log(`Can Not Remove Files! ${filePath}`);}
            }
        }); 
        try{
            fs.rmdirSync(dirPath);
        }
        catch(err){console.log(`Can Not Remove Directory! ${dirPath}`);}
    }
}
function flatJson(json){
    if(typeof json === 'string') json = JSON.parse(json);
    let json2 = [];
    function traverse(item){
        if(typeof item === 'object'){
            if('length' in item){
                for(let i of item){
                    traverse(i);
                }
            }
            else{
                for(let k in item){
                    json2.push(k);
                    traverse(item[k]);
                }
            }
        }
        else json2.push(item);
    }
    traverse(json);
    return(json2);
}
let dataPath = 'data';
function jsonData(name){
    let content = fs.readFileSync(path.join(dataPath, `json/${name}.json`), 'utf-8');
    let json = JSON.parse(content);
    return(json);
}
function pick(arr){
    return(arr[Math.floor(Math.random()*arr.length)]);
}
String.prototype.toFullShape = function(){
    let text = this.valueOf();
    let list = text.split('');
    list = list.map(char => {
        let charCode = char.charCodeAt(0);
        if(charCode >= 32 && charCode < 32+95){
            if(char == ' ') return('　');
            return(String.fromCharCode(charCode - 32 + 65280));
        }
        return(char);
    });
    return(list.join(''));
}
String.prototype.toHalfShape = function(){
    let text = this.valueOf();
    let list = text.split('');
    list = list.map(char => {
        let charCode = char.charCodeAt(0);
        if(charCode >= 65280 && charCode < 65280+95){
            if(char == '　') return(' ');
            return(String.fromCharCode(charCode - 65280 + 32));
        }
        return(char);
    });
    return(list.join(''));
}
function alert(text){
    dialog.showMessageBox({message: text});
}

// docx method
let docxPath = '';
let unpackedPath = 'unpacked';
let docxExtractedData = {};
function selectDocx(){
    return dialog.showOpenDialog({
        title: 'AutoPPT', 
        // buttonLabel: 'Open', 
        filters: [
            { name: 'Word', extensions: ['docx', 'doc'] },
            { name: 'Zip', extensions: ['zip'] }, 
            { name: 'All', extensions: ['*'] }
        ], 
        properties: ['openFile', 'treatPackageAsDirectory'], 
        message: 'select a ".docx" file'
    })
    .then(data => data?.filePaths[0])
    .catch(err => {console.log(err);});
}
function unpackDocx(filePath){
    deleteDir(unpackedPath);
    return decompress(filePath, unpackedPath)
    .catch((err) => {console.log(err);});
}
function relsDocx(dirPath){
    const parser = new XMLParser({
        ignoreAttributes : false
    });
    let relsPath = path.join(dirPath, '_rels/.rels');
    let documentPath = path.join(dirPath, '/word/document.xml');
    let dp2drp = path => {
        var usingBackslash = path.indexOf('\\') > -1;
        let pathList = path.replaceAll('\\', '/').split('/');
        let last = pathList.pop();
        pathList.push('_rels');
        pathList.push(last);
        path = pathList.join('/');
        path += '.rels';
        if(usingBackslash) path.replaceAll('/', '\\');
        return(path);
    }
    let documentRelsPath = dp2drp(documentPath);
    let embedRels = {};
    if(fs.existsSync(relsPath)){
        data = fs.readFileSync(relsPath, 'utf-8')
        let rels = parser.parse(data);
        var targetPath = rels?.Relationships?.Relationship['@_Target'];
        if(targetPath){
            documentPath = path.join(dirPath, targetPath);
            documentRelsPath = dp2drp(documentPath);
        }
    }
    if(fs.existsSync(documentRelsPath)){
        data = fs.readFileSync(documentRelsPath, 'utf-8')
        let documentRels = parser.parse(data);
        let embedDirPath = path.dirname(path.dirname(documentRelsPath));
        for(let r of documentRels?.Relationships?.Relationship){
            embedRels[r['@_Id']] = path.join(embedDirPath, r['@_Target']);
        }
    }
    if(fs.existsSync(documentPath)){
        data = fs.readFileSync(documentPath, 'utf-8')
        const parser = new XMLParser({
            ignoreAttributes : false
        });
        let document = data;
        // fs.writeFileSync(documentPath+'.json', JSON.stringify(document, true, 4));
        return({documentPath, document, embedRels});
    }
    return({documentPath, document: '', embedRels});
}
// function convertDocx(dirPath){
//     const parser = new XMLParser({
//         ignoreAttributes : false
//     });
//     let relsPath = path.join(dirPath, '_rels/.rels');
//     let documentPath = path.join(dirPath, '/word/document.xml');
//     let dp2drp = path => {
//         var usingBackslash = path.indexOf('\\') > -1;
//         let pathList = path.replaceAll('\\', '/').split('/');
//         let last = pathList.pop();
//         pathList.push('_rels');
//         pathList.push(last);
//         path = pathList.join('/');
//         path += '.rels';
//         if(usingBackslash) path.replaceAll('/', '\\');
//         return(path);
//     }
//     let documentRelsPath = dp2drp(documentPath);
//     let embedRels = {};
//     if(fs.existsSync(relsPath)){
//         data = fs.readFileSync(relsPath, 'utf-8')
//         let rels = parser.parse(data);
//         var targetPath = rels?.Relationships?.Relationship['@_Target'];
//         if(targetPath){
//             documentPath = path.join(dirPath, targetPath);
//             documentRelsPath = dp2drp(documentPath);
//         }
//     }
//     if(fs.existsSync(documentRelsPath)){
//         data = fs.readFileSync(documentRelsPath, 'utf-8')
//         let documentRels = parser.parse(data);
//         let embedDirPath = path.dirname(path.dirname(documentRelsPath));
//         for(let r of documentRels?.Relationships?.Relationship){
//             embedRels[r['@_Id']] = path.join(embedDirPath, r['@_Target']);
//         }
//     }
//     if(fs.existsSync(documentPath)){
//         data = fs.readFileSync(documentPath, 'utf-8')
//         const parser = new XMLParser({
//             ignoreAttributes : false
//         });
//         let document = parser.parse(data);
//         // fs.writeFileSync(documentPath+'.json', JSON.stringify(document, true, 4));
//         // fs.writeFileSync(documentPath+'.txt', flatJson(document).join('\n'));
//         return({documentPath, document, embedRels});
//     }
//     return({documentPath, document: [], embedRels});
// }
// function extractDocx(json, embedRels){
//     json = flatJson(json);
//     let outputFlag = false;
//     let embedFlag = true;
//     let tableFlag = false;
//     let tableNameFlag = false;
//     let data = {};
//     let pages = [[]];
//     let createTable = () => [];
//     let tableNameNow = '';
//     let tables = {};
//     let tableAttr = ['w:tblStyle', '@_w:val', 'w:tblW', '@_w:w', '@_w:type', 'w:jc', 'w:tblBorders', 'w:top', '@_w:color', '@_w:space', '@_w:sz', 'w:left', 'w:bottom', 'w:right', 'w:insideH', 'w:insideV', 'w:tblLayout', 'w:tblLook', 'w:tblGrid'];
//     let paraIdFlag = false;
//     let paraIdNow = false;
//     for(let row of json){
//         if(tableNameFlag){
//             if(tableAttr.indexOf(row) == -1){
//                 tableNameNow = row;
//                 tables[row] = createTable();
//                 tableNameFlag = false;
//             }
//         }
//         let tableNow = tables[tableNameNow];
//         if(outputFlag || embedFlag){
//             let rowOri = row;
//             row = embedFlag ? `embed(${embedRels[row] || row})` : row;
//             if(tableFlag){
//                 gridNow = tableNow[tableNow.length-1];
//                 gridNow[gridNow.length-1].push(row);
//             }
//             else{
//                 data[data.length-1].push(row);
//             }
//             if(outputFlag) outputFlag = false;
//             if(embedFlag) embedFlag = false;
//             row = rowOri;
//         }
//         switch(row){
//             case '#text':
//                 outputFlag = true;
//                 break;
//             case '@_r:embed':
//                 embedFlag = true;
//                 break;
//             case 'w:tbl':
//                 tableFlag = true;
//                 break;
//             case 'w:tblPr':
//                 tableNameFlag = true;
//                 break;
//             case 'w:trPr':
//                 tableNow.push([]);
//                 break;
//             case 'w:tcPr':
//                 tableNow[tableNow.length-1].push([]);
//                 break;
//         }
//     }
//     // data = data.filter(page => page.length > 0);
//     // console.log(data);
//     // for(let name in tables){
//     //     console.table(tables[name].map(i => i.map(j => j.join('\n'))));
//     // }
//     return({data, tables});
// }
function extractDocx2(data){
    let doc = HTMLparser(data.document);
    rData = [];
    let convertP = element => {
        let embed = $('a\\:blip', element)?.getAttribute('r:embed');
        return({
            id: element.getAttribute('w14:paraId'), 
            type: embed !== undefined ? 'embed' : $('w\\:pPr > w\\:pStyle', element)?.getAttribute('w:val'), 
            text: $('w\\:r > w\\:t', element)?.innerText, 
            embed: data.embedRels[embed] || embed
        });
    }
    let convertTable = element => {
        return({
            id: $('w\\:tblPr > w\\:tblStyle', element)?.getAttribute('w:val'), 
            type: 'table', 
            grid: [...$$('w\\:tr', element)]?.map(tr => 
                    [...$$('w\\:tc', tr)]?.map(tc => 
                        [...$$('w\\:p', tc)]?.map(p => 
                            convertP(p)
                        )
                    )
                )
        });
    }
    $$('w\\:document > w\\:body > *', doc).forEach(element => {
        switch(element.tagName.toLowerCase()){
            case 'w:p':
                rData.push(convertP(element));
                break;
            case 'w:tbl':
                rData.push(convertTable(element));
                break;
        }
    });
    // fs.writeFileSync('test.json', JSON.stringify(rData, true, 4));
    var typeIndexTable = {
        title: 0, 
        subtitle: 1, 
        heading1: 2, 
        heading2: 3, 
        heading3: 4, 
        heading4: 5, 
        heading5: 6, 
        heading6: 7
    };
    let typeList = Object.keys(typeIndexTable);
    let typeIndexList = [];
    rData.map(item => {
        let typeIndex = item.type !== undefined ? typeIndexTable[item.type.toLowerCase()] : undefined;
        if(typeIndex !== undefined){
            item.typeIndex = typeIndex;
            if(typeIndexList.indexOf(typeIndex) === -1) typeIndexList.push(typeIndex);
        }
    });
    typeIndexList = typeIndexList.sort((a, b) => a - b);
    rData.map(item => {
        if(item.typeIndex !== undefined){
            let newType = typeList[typeIndexList.indexOf(item.typeIndex)];
            if(newType !== undefined) item.type = newType;
            delete item.typeIndex;
        }
    });
    return(rData);
}
function extractedData2html(data/*extract2*/){
    html = '';
    function item2html(item){
        let tempHtml = '';
        switch(item.type){
            case 'table':
                var innerHtml = item.grid.map(row => `<tr>${row.map(cell => `<td>${cell.map(subItem => item2html(subItem)).join('')}</td>`).join('')}</tr>`).join('');
                tempHtml = `<table data-id="${item.id}">${innerHtml}</table>`;
                break;
            case 'embed':
                tempHtml = `<img src="../${item.embed}" alt="${item.id}" data-id="${item.id}"></img>`;
                break;
            default:
                var contrastTable = {
                    title: 'h1', 
                    subtitle: 'h2', 
                    heading1: 'h3', 
                    heading2: 'h4', 
                    heading3: 'h4', 
                    heading4: 'h5', 
                    heading5: 'h5', 
                    heading6: 'h6'
                };
                var tagName = contrastTable[item.type?.toLowerCase()] || 'p';
                tempHtml = item.text !== undefined && item.text?.length > 0 ? `<${tagName} data-id="${item.id}">${item.text}</${tagName}>` : `<br data-id="${item.id}>`;
                break;
        }
        return(tempHtml);
    }
    for(let item of data){
        html += item2html(item);
    }
    return(html);
}
function analyzeText(text){
    let analyzedData = {};
    let subjectTypesValue = {};
    let subjectTypes = jsonData('subjectTypes');
    subjectTypes.forEach(type => {
        subjectTypesValue[type] = 0;
    });
    let subjectKeywords = jsonData('subjectKeywords');
    for(let keyword in subjectKeywords){
        let flag = false;
        if(text.indexOf(keyword) > -1) flag = true;
        for(let kw of subjectKeywords[keyword][0]){
            if(text.indexOf(kw) > -1) flag = true;
        }
        if(flag){
            for(let type in subjectKeywords[keyword][1]){
                if(subjectTypesValue[type] !== undefined) subjectTypesValue[type] += subjectKeywords[keyword][1][type];
            }
        }
    }
    analyzedData.subject = Object.entries(subjectTypesValue).sort((i1, i2) => i2[1] - i1[1])[0][0];
    console.log(analyzedData.subject);
    return(analyzedData);
}
function uploadDocx(event){
    let errMsg = {
        noDocxFile: 'hasn\'t select any docx file'
    }
    return selectDocx()
    .then(filePath => {
        docxPath = filePath ? filePath : '';
        if(filePath) return(unpackDocx(filePath));
        else throw new Error(errMsg.noDocxFile);
    })
    // .then(unpackedData => convertDocx(unpackedPath))
    // .then(data => extractDocx(data['document'], data['embedRels']));
    .then(unpackedData => relsDocx(unpackedPath))
    .then(data => extractDocx2(data))
    .then(data => {
        docxExtractedData = data;
        let html = extractedData2html(data);
        let analyzedData = analyzeText(HTMLparser(html).innerText);
        let settings = jsonData('settingsForm');
        let subjectTypes = jsonData('subjectTypes');
        settings.subject['type'] = settings.subject['type'].replace('()', `(${subjectTypes.join(', ')})`);
        Object.entries(analyzedData).forEach(kv => {
            settings[kv[0]].value = kv[1];
        });
        return({settings, html});
    })
    .catch(err => {
        if(err.message !== errMsg.noDocxFile) console.log(err);
    });
}

// pptx method
async function generatePptx(settings){
    // data
    let layouts = {
        LAYOUT_4x3: { type: 'screen4x3', width: 9144000, height: 6858000 },
        LAYOUT_16x9: { type: 'screen16x9', width: 9144000, height: 5143500 },
        LAYOUT_16x10: { type: 'screen16x10', width: 9144000, height: 5715000 },
        LAYOUT_WIDE: { type: 'custom', width: 12191996, height: 6858000 },
        LAYOUT_USER: { type: 'custom', width: 12191996, height: 6858000 },
    };
    let lengthRatio = 1/1e4/32.24*25.4;
    let slideWidth = layouts[settings.layout.value].width*lengthRatio;
    let slideHeight = layouts[settings.layout.value].height*lengthRatio;

    // image
    function getImageData(path){
        let name = path.replaceAll('\\', '/').split('/').pop();
        let list = name.split('.');
        var extension = list.pop();
        list = list.join('.').split('_');
        list.push(extension);
        'game_background_1_dc_neo'
        if(list.length < 6){
            throw new Error(`Can Not Get Image Data! (${path})`);
        }
        else{
            return({
                subject: list[0], 
                type: list[1], 
                num: list[2], 
                BorD: {b: 'bright', d: 'dark'}[list[3][0]], 
                focalPoint: {t: 'top', b: 'bottom', l: 'left', r: 'right', c: 'center', s: 'surround'}[list[3][1]], 
                theme: list[4]
            });
        }
    }
    let imageDirPath = path.join(dataPath, 'image');
    let images = fs.readdirSync(imageDirPath);
    let usableImages = images.filter(name => name.indexOf(settings.subject.value) == 0);
    let bgis = usableImages.filter(name => name.indexOf('background') == settings.subject.value.length+1);
    let bgiPath = path.join(imageDirPath, pick(bgis));
    let bgiData = getImageData(bgiPath);
    function addBgi(slide, bgiPath){
        let imageSize = getImageSize(bgiPath);
        let imageAlign = imageSize.width/imageSize.heigth > slideWidth/slideHeight ? 'height' : 'width';
        let cx = parseInt(imageAlign == 'width' ? slideWidth : slideHeight/imageSize.height*imageSize.width);
        let cy = parseInt(imageAlign == 'height' ? slideHeight : slideWidth/imageSize.width*imageSize.height);
        let x = -(cx-slideWidth)/2;
        let y = -(cy-slideHeight)/2;
        slide.addImage({
            file: bgiPath, 
            x, y, cx, cy
        });
    }

    // attribute
    let fontFace = 'Gen Jyuu Gothic Bold';
    let fontSize = 20;
    let gap = 5;

    // pptx
    const PPTX = require('nodejs-pptx');
    let pptx = new PPTX.Composer();
    function titlePage(pageItems){
        return slide => {
            addBgi(slide, bgiPath);
            let title = pageItems.shift();
            let titleSize = 30;
            let presetType = 'center';
            switch(bgiData.focalPoint){
                case 'center':
                case 'surround':
                    presetType = 'center';
                    break;
                case 'left':
                    presetType = 'right';
                    break;
                case 'right':
                    presetType = 'left';
                    break;
                case 'top':
                    presetType = 'bottom';
                    break;
                case 'bottom':
                    presetType = 'top';
                    break;
            }
            presetType = 'bottom';
            let shapePreset = {
                center: {
                    x: (slideWidth - slideWidth/6*5)/2, 
                    y: (slideHeight - slideHeight/2)/2, 
                    cx: slideWidth/6*5, 
                    cy: slideHeight/2, 
                }, 
                top: {
                    x: 0, 
                    y: 0, 
                    cx: slideWidth, 
                    cy: slideHeight/2, 
                }, 
                bottom: {
                    x: 0, 
                    y: slideHeight - slideHeight/2, 
                    cx: slideWidth, 
                    cy: slideHeight/2, 
                }, 
                left: {
                    x: 0, 
                    y: 0, 
                    cx: slideWidth/3, 
                    cy: slideHeight, 
                }, 
                right: {
                    x: slideWidth - slideWidth/3, 
                    y: 0, 
                    cx: slideWidth/3, 
                    cy: slideHeight, 
                }
            };
            let titlePreset = {
                center: {
                    x: (slideWidth - slideWidth/2)/2, 
                    y: (slideHeight - titleSize)/2, 
                    cx: slideWidth/2, 
                    textAlign: 'center'
                }, 
                top: {
                    x: 20, 
                    y: slideHeight/4 - titleSize/2, 
                    cx: slideWidth/2, 
                    textAlign: 'left'
                }, 
                bottom: {
                    x: slideWidth - 20 - slideWidth/2, 
                    y: slideHeight/4*3 - titleSize/2, 
                    cx: slideWidth/2, 
                    textAlign: 'right'
                }, 
                left: {
                    x: (slideWidth/3 - slideWidth/4)/2, 
                    y: (slideHeight - titleSize)/2, 
                    cx: slideWidth/4, 
                    textAlign: 'center'
                }, 
                right: {
                    x: slideWidth/3*2 + (slideWidth/3 - slideWidth/4)/2, 
                    y: (slideHeight - titleSize)/2, 
                    cx: slideWidth/4, 
                    textAlign: 'center'
                }
            };
            let cx = titlePreset[presetType].cx;
            if(cx < title.text.length * titleSize){
                if(/( |　)/.test(title.text)){
                    let rows = [];
                    let rowNow = '';
                    let list = title.text.replaceAll('　', ' ').split(' ');
                    for(let i = 0; i < list.length; i++){
                        rowNow += list[i]+' ';
                        if(i == list.length-1){
                            rows.push(rowNow);
                            rowNow = '';
                        }
                        // if(rowNow.length * titleSize < cx){
                            if((rowNow.length+list[i+1]) * titleSize > cx){
                                rows.push(rowNow);
                                rowNow = '';
                            }
                        // }
                    }
                    title.text = rows.join('\n');
                }
                else{
                    let rowNum = title.text.length*titleSize / cx;
                    let rowLength = Math.round(title.text.length/rowNum);
                    let rows = [];
                    let rowNow = '';
                    for(let i = 0; i < title.text.length; i++){
                        rowNow += title.text[i];
                        if(i % rowLength == rowLength-1 || i == title.text.length-1){
                            rows.push(rowNow);
                            rowNow = '';
                        }
                    }
                    title.text = rows.join('\n');
                }
            }
            slide.addShape({
                type: PPTX.ShapeTypes.RECTANGLE, 
                color: '888888', 
                ...shapePreset[presetType]
            });
            slide.addText({
                value: title.text.toFullShape(), 
                fontFace: fontFace, 
                fontSize: titleSize, 
                textColor: '000000', 
                textWrap: 'none', 
                textVerticalAlign: 'center', 
                margin: 0, 
                ...titlePreset[presetType]
            });
            pageItems.map((item, i) => {
                slide.addText({
                    value: item.text, 
                    x: gap, 
                    y: (fontSize+gap) * i, 
                    fontFace: fontFace, 
                    fontSize: fontSize, 
                    textColor: '000000', 
                    textWrap: 'none', 
                    textAlign: 'left', 
                    textVerticalAlign: 'center', 
                    margin: 0
                });
            });
        }
    }
    function defaultPage(pageItems){
        return slide => {
            addBgi(slide, bgiPath);
            pageItems.map((item, i) => {
                slide.addText({
                    value: item.text, 
                    x: gap, 
                    y: (fontSize+gap) * i, 
                    fontFace: fontFace, 
                    fontSize: fontSize, 
                    textColor: '000000', 
                    textWrap: 'none', 
                    textAlign: 'left', 
                    textVerticalAlign: 'center', 
                    // line: { color: '0000FF', dashType: 'dash', width: 1.0 }, 
                    margin: 0
                });
            });
        }
    }
    await pptx.compose(pres => {
        pres
        .title(settings.title.value)
        .author(settings.author.value)
        .company('MaoHuPi - Auto PPT')
        .revision(settings.revision.value)
        .subject(settings.subject.value)
        .layout(settings.layout.value);
        let pagesData = [];
        for(let item of docxExtractedData){
            if(['title', 'subtitle', 'heading1'].indexOf(item.type?.toLowerCase()) > -1) pagesData.push([]);
            if(pagesData[pagesData.length-1] === undefined) pagesData.push([]);
            pagesData[pagesData.length-1].push(item);
        }
        for(let pageItems of pagesData){
            if(pageItems[0].type.toLowerCase() == 'title'){
                pres.addSlide(titlePage(pageItems));
            }
            else{
                pres.addSlide(defaultPage(pageItems));
            }
        }
    });
    return(pptx);
}
async function exportPptx(event, settings){
    let outputPath = `unpacked/output.pptx`;
    let pptx = await generatePptx(settings);
    await pptx.save(outputPath);
    return(outputPath);
}

// electron method
let  iconPath = path.join(__dirname, 'web/image/logo.png')
function createWindow(){
    const appIcon = new Tray(iconPath);
    const win = new BrowserWindow({
        width: 800, 
        height: 600, 
        webPreferences: {
            preload: path.join(__dirname, 'preload.js')
        }, 
        icon: iconPath
    });

    win.setMenuBarVisibility(false);
    win.setResizable(true);
    win.setMinimumSize(400, 400);
    win.setMinimizable(true);
    win.setMaximizable(true);

    win.loadFile('web/index.html');
}

app.whenReady().then(() => {
    // ipcMain.on('uploadDocx', uploadDocx);
    ipcMain.handle('uploadDocx', uploadDocx);
    ipcMain.handle('exportPptx', exportPptx);

    createWindow();

    app.on('activate', () => {
        if (BrowserWindow.getAllWindows().length === 0) {
            createWindow();
        }
    })
})

app.on('window-all-closed', () => {
    if (process.platform !== 'darwin') {
        app.quit();
    }
});