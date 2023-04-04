// basic import
const path = require('path');
const { app, BrowserWindow, Tray, ipcMain } = require('electron');

// module import
const fs = require('fs');
const { dialog } = require('electron');
const decompress = require('decompress');
const { parse } = require('node-html-parser');
const { XMLParser, XMLBuilder, XMLValidator} = require('fast-xml-parser');

// basic method
let $ = (e, p = document) => p.querySelector(e);
let $$ = (e, p = document) => p.querySelectorAll(e);
function deleteDir(dirPath){ 
    if(fs.existsSync(dirPath)){
        let filePathList = fs.readdirSync(dirPath) || [];
        filePathList.forEach(filePath => {
            filePath = path.join(dirPath, filePath); 
            if(fs.statSync(filePath).isDirectory()) deleteDir(filePath);
            else fs.unlinkSync(filePath);
        }); 
        fs.rmdirSync(dirPath);
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

// docx method
let docxPath = '';
let unpackedPath = 'unpacked';
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
// function analyzeDocx(json, embedRels){
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
function analyzeDocx2(data){
    let doc = parse(data.document);
    rData = [];
    let convertP = element => {
        let embed = $('a\\:blip', element)?.getAttribute('r:embed');
        return({
            id: element.getAttribute('w14:paraId'), 
            type: $('w\\:pPr > w\\:pStyle', element)?.getAttribute('w:val') || (embed !== undefined ? 'embed' : undefined), 
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
    return(rData);
}
function uploadDocx(event){
    return selectDocx()
    .then(filePath => {
        docxPath = filePath ? filePath : '';
        if(filePath){
            return(unpackDocx(filePath));
        }
    })
    // .then(unpackedData => convertDocx(unpackedPath))
    // .then(data => analyzeDocx(data['document'], data['embedRels']));
    .then(unpackedData => relsDocx(unpackedPath))
    .then(data => analyzeDocx2(data));
}

// pptx method

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