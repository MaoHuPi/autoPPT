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
const JSZip = require('jszip');
const xml2js = require('xml2js');

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
function textData(name){
    let content = fs.readFileSync(path.join(dataPath, `text/${name}.txt`), 'utf-8');
    return(content);
}
function tsvData(name){
    let splitWith = '\t';
    let content = fs.readFileSync(path.join(dataPath, `tsv/${name}.tsv`), 'utf-8');
    let rows = content.replaceAll('\r', '').split('\n');
    let columnTitle = rows.shift().split(splitWith);
    rows = rows.map(text => Object.fromEntries(text.split(splitWith).map((value, i) => [columnTitle[i], value])));
    return(rows);
}
function pick(arr){
    return(arr[Math.floor(Math.random()*arr.length)]);
}
const fullShapeSpace = '　';
String.prototype.toFullShape = function(){
    let text = this.valueOf();
    let list = text.split('');
    list = list.map(char => {
        let charCode = char.charCodeAt(0);
        if(charCode >= 32 && charCode < 32+95){
            if(char == ' ') return(fullShapeSpace);
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
            if(char == fullShapeSpace) return(' ');
            return(String.fromCharCode(charCode - 65280 + 32));
        }
        return(char);
    });
    return(list.join(''));
}
function alert(text){
    dialog.showMessageBox({message: text});
}

// basic data
let subjectTypes = textData('subjectTypes').replaceAll('\r', '').split('\n');

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
    let stylesPath = path.join(dirPath, '/word/styles.xml');
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
    let styleTable = {};
    if(fs.existsSync(relsPath)){
        data = fs.readFileSync(relsPath, 'utf-8');
        let rels = parser.parse(data);
        var targetPath = rels?.Relationships?.Relationship['@_Target'];
        if(targetPath){
            documentPath = path.join(dirPath, targetPath);
            documentRelsPath = dp2drp(documentPath);
        }
    }
    if(fs.existsSync(documentRelsPath)){
        data = fs.readFileSync(documentRelsPath, 'utf-8');
        let documentRels = parser.parse(data);
        let embedDirPath = path.dirname(path.dirname(documentRelsPath));
        for(let r of documentRels?.Relationships?.Relationship){
            embedRels[r['@_Id']] = path.join(embedDirPath, r['@_Target']);
        }
    }
    if(fs.existsSync(stylesPath)){
        data = fs.readFileSync(stylesPath, 'utf-8');
        let styles = parser.parse(data);
        for(let style of styles['w:styles']['w:style']){
            if(style['w:name']) styleTable[style['@_w:styleId']] = style['w:name']['@_w:val'].replaceAll(' ', '').toLowerCase();
        }
    }
    if(fs.existsSync(documentPath)){
        data = fs.readFileSync(documentPath, 'utf-8');
        let document = data;
        // fs.writeFileSync(documentPath+'.json', JSON.stringify(document, true, 4));
        return({documentPath, document, embedRels, styleTable});
    }
    return({documentPath, document: '', embedRels, styleTable});
}
function extractDocx(data){
    let doc = HTMLparser(data.document);
    rData = [];
    let convertP = element => {
        let embed = $('a\\:blip', element)?.getAttribute('r:embed');
        let type = embed !== undefined ? 'embed' : $('w\\:pPr > w\\:pStyle', element)?.getAttribute('w:val');
        type = data.styleTable[type] !== undefined ? data.styleTable[type] : type;
        return({
            id: element.getAttribute('w14:paraId'), 
            type: type, 
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
                tempHtml = item.text !== undefined && item.text?.length > 0 ? `<${tagName} data-id="${item.id}">${item.text}</${tagName}>` : `<br data-id="${item.id}">`;
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
    .then(unpackedData => relsDocx(unpackedPath))
    .then(data => extractDocx(data))
    .then(data => {
        docxExtractedData = data;
        let html = extractedData2html(data);
        let analyzedData = analyzeText(HTMLparser(html).innerText);
        let settings = {};
        tsvData('settingsForm').forEach(item => {
            settings[item.key] = {...item, key: undefined};
        });
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

    // method
    function fillInImage(rect = {x: 0, y: 0, cx: 10, cy: 10}, imagePath = '', mode = ''/* internal, overflow, stretch */){
        if(mode == 'stretch') return(rect);
        let imageSize = getImageSize(imagePath);
        let imageAlign = imageSize.width/imageSize.heigth > rect.cx/rect.cy ? 'height' : 'width';
        var cxOptions = [rect.cx, rect.cy/imageSize.height*imageSize.width];
        var cyOptions = [rect.cx, rect.cy/imageSize.height*imageSize.width];
        var cx, cy, x, y;
        cx = parseInt(imageAlign === 'width' ^ mode !== 'overflow' ? cxOptions[0] : cxOptions[1]);
        cy = parseInt(imageAlign === 'height' ^ mode !== 'overflow' ? cyOptions[0] : cyOptions[1]);
        x = -(cx-rect.cx)/2;
        y = -(cy-rect.cy)/2;
        return({x, y, cx, cy});
    }
    function fillInText(rect = {x: 0, y: 0, cx: 10, cy: 10}, text = '', fontSize = 20){
        let cx = rect.cx;
        if(cx < text.length * fontSize){
            if(new RegExp(`( |${fullShapeSpace})`).test(text)){
                let rows = [];
                let rowNow = '';
                let list = text.replaceAll(fullShapeSpace, ' ').split(' ');
                for(let i = 0; i < list.length; i++){
                    rowNow += list[i]+' ';
                    if(i == list.length-1){
                        rows.push(removeStartEndSpace(rowNow));
                        rowNow = '';
                    }
                    // if(rowNow.length * fontSize < cx){
                        if(list[i+1] !== undefined && (rowNow.length+list[i+1].length) * fontSize > cx){
                            rows.push(removeStartEndSpace(rowNow));
                            rowNow = '';
                        }
                    // }
                }
                text = rows.join('\n');
            }
            else{
                let rowNum = text.length*fontSize / cx;
                let rowLength = Math.round(text.length/rowNum);
                let rows = [];
                let rowNow = '';
                for(let i = 0; i < text.length; i++){
                    rowNow += text[i];
                    if(i % rowLength == rowLength-1 || i == text.length-1){
                        rows.push(rowNow);
                        rowNow = '';
                    }
                }
                text = rows.join('\n');
            }
        }
        let newRect = {};
        newRect.cx = rect.cx;
        newRect.cy = text.split('\n').length * fontSize;
        newRect.x = rect.x;
        rect.cy = rect.cy || fontSize;
        newRect.y = rect.y - (newRect.cy - rect.cy)/2;
        return({rect: newRect, text});
    }
    function resetElementCNvPr(slide, nameList){
        let spTree = slide.powerPointFactory.pptFactory.slideFactory.content[`ppt/slides/${slide.name}.xml`]['p:sld']['p:cSld'][0]['p:spTree'][0];
        let nthElementCount = {'p:sp': 1};
        let tagNametable = {
            TextBox: ['p:sp', 'p:nvSpPr'], 
            Shape: ['p:sp', 'p:nvSpPr'], 
            Image: ['p:pic', 'p:nvPicPr']
        };
        slide.elements.map((element, i) => {
            let tnd/*tagNameData*/ = tagNametable[element.constructor.name];
            if(nthElementCount[tnd[0]] === undefined) nthElementCount[tnd[0]] = 0;
            spTree[tnd[0]][nthElementCount[tnd[0]]][tnd[1]][0]['p:cNvPr'][0].$.id = i+1;
            spTree[tnd[0]][nthElementCount[tnd[0]]][tnd[1]][0]['p:cNvPr'][0].$.name = nameList[i];
            nthElementCount[tnd[0]]++;
        });
    }
    function removeStartEndSpace(text){
        return(text.replaceAll(/(^( |\t|\s)+|( |\t|\s)+$)/g, ''));
    }
    function zipContent(){
        /* Override the function to separate the parts to be sorted. */
        function createTag(xmlObj){
            return(xmlObj ? Object.keys(xmlObj).filter(key => key !== '$').map(tagName => `<${tagName}${xmlObj[tagName][0]?.$ !== undefined ? ' '+Object.entries(xmlObj[tagName][0]?.$).map(kv => `${kv[0]}="${kv[1]}"`).join(' ') : ''}>${createTag(xmlObj[tagName][0])}</${tagName}>`).join('') : '');
        }
        let zip = new JSZip();
        let content = this.content;
        for (let key in content) {
            if (content.hasOwnProperty(key)) {
                let ext = key.substr(key.lastIndexOf('.'));
                if (ext === '.xml' || ext === '.rels') {
                    let builder = new xml2js.Builder({ renderOpts: { pretty: false } });
                    let $$, $$_xml;
                    if(content[key].temp$$ !== undefined){
                        $$ = content[key].temp$$;
                        delete content[key].temp$$;
                        $$_xml = $$
                        .map(d => {
                            let n = d['#element'];
                            let xmlContent;
                            try{
                                let xml = builder.buildObject(n);
                                xmlContent = /<root>(.*)<\/root>/.exec(xml)[1];
                            }
                            catch(err){
                                // console.log(err);
                                xmlContent = createTag(n);
                            }
                            return(`<${d['#name']}>${xmlContent}</${d['#name']}>`);
                        })
                        .join('');
                    }
                    let xml = builder.buildObject(content[key]);
                    if($$_xml !== undefined){
                        xml = xml.replace(/<p:spTree>.*<\/p:spTree>/g, `<p:spTree>${$$_xml}</p:spTree>`);
                    };
                    zip.file(key, xml);
                } else {
                    zip.file(key, content[key]);
                }
            }
        }
        return zip;
    }

    // image
    let imageDirPath = path.join(dataPath, 'image');
    let bgiDf = tsvData('bgiData');
    let bgis = bgiDf.filter(bgi => bgi.subject == settings.subject.value);
    let bgiData = pick(bgis);
    let bgiPath = path.join(imageDirPath, bgiData.filePath);
    function addBgi(slide, bgiPath){
        let newRect = fillInImage({x: 0, y: 0, cx: slideWidth, cy: slideHeight}, bgiPath, 'overflow');
        slide.addImage({
            file: bgiPath, 
            ...newRect
        });
    }

    // attribute
    // let fontFace = 'Gen Jyuu Gothic Bold';
    let fontFace = '微軟正黑體';
    let fontSize = 20;
    let gap = 5;

    // pptx
    const PPTX = require('nodejs-pptx');
    let pptx = new PPTX.Composer();
    function titlePage(pageItems){
        return slide => {
            let title = pageItems.shift();
            let titleSize = 30;
            let presetType = 'center';
            switch(bgiData.focalPoint){
                case 'c':
                case 's':
                    presetType = 'center';
                    break;
                case 'l':
                    presetType = 'right';
                    break;
                case 'r':
                    presetType = 'left';
                    break;
                case 't':
                    presetType = 'bottom';
                    break;
                case 'b':
                    presetType = 'top';
                    break;
            }
            presetType = 'left';
            let preset = {
                center: [
                    {
                        shape: {
                            x: (slideWidth - slideWidth/6*5)/2, 
                            y: (slideHeight - slideHeight/2)/2, 
                            cx: slideWidth/6*5, 
                            cy: slideHeight/2, 
                        }, 
                        title: {
                            x: (slideWidth - slideWidth/2)/2, 
                            y: (slideHeight - titleSize)/2, 
                            cx: slideWidth/2, 
                            textAlign: 'center'
                        }, 
                        descriptions: {
                            x: 40, 
                            y: slideHeight/4*3 + slideHeight/4/2 - fontSize/2, 
                            cx: slideWidth - 40*2, 
                            textAlign: 'center'
                        }
                    }
                ], 
                top: [
                    {
                        shape: {
                            x: 0, 
                            y: 0, 
                            cx: slideWidth, 
                            cy: slideHeight/2, 
                        }, 
                        title: {
                            x: 20, 
                            y: slideHeight/4 - titleSize/2, 
                            cx: slideWidth/2, 
                            textAlign: 'left'
                        }, 
                        descriptions: {
                            x: 40, 
                            y: slideHeight/2 + slideHeight/2/2 - fontSize/2, 
                            cx: slideWidth - 40*2, 
                            textAlign: 'center'
                        }
                    }
                ], 
                bottom: [
                    {
                        shape: {
                            x: 0, 
                            y: slideHeight - slideHeight/2, 
                            cx: slideWidth, 
                            cy: slideHeight/2, 
                        }, 
                        title: {
                            x: slideWidth - 20 - slideWidth/2, 
                            y: slideHeight/4*3 - titleSize/2, 
                            cx: slideWidth/2, 
                            textAlign: 'right'
                        }, 
                        descriptions: {
                            x: 40, 
                            y: slideHeight/2/2 - fontSize/2, 
                            cx: slideWidth - 40*2, 
                            textAlign: 'center'
                        }
                    }
                ], 
                left: [
                    {
                        shape: {
                            x: 0, 
                            y: 0, 
                            cx: slideWidth/3, 
                            cy: slideHeight, 
                        }, 
                        title: {
                            x: (slideWidth/3 - slideWidth/4)/2, 
                            y: (slideHeight - titleSize)/2, 
                            cx: slideWidth/4, 
                            textAlign: 'center'
                        }, 
                        descriptions: {
                            x: slideWidth/3 + 40, 
                            y: slideHeight/2 - fontSize/2, 
                            cx: slideWidth/3*2 - 40*2, 
                            textAlign: 'center'
                        }
                    }
                ], 
                right: [
                    {
                        shape: {
                            x: slideWidth - slideWidth/3, 
                            y: 0, 
                            cx: slideWidth/3, 
                            cy: slideHeight, 
                        }, 
                        title: {
                            x: slideWidth/3*2 + (slideWidth/3 - slideWidth/4)/2, 
                            y: (slideHeight - titleSize)/2, 
                            cx: slideWidth/4, 
                            textAlign: 'center'
                        }, 
                        descriptions: {
                            x: 40, 
                            y: slideHeight/2 - fontSize/2, 
                            cx: slideWidth/3*2 - 40*2, 
                            textAlign: 'center'
                        }
                    }
                ]
            };
            let titleGap = 20;
            let usingPreset = pick(preset[presetType]);
            title.text = removeStartEndSpace(title.text);
            let titleFillInData = fillInText({...usingPreset.title}, title.text, titleSize);
            let nameList = [];
            title.text = titleFillInData.text;
            let subTitle, subTitleSize, subTitleFillInData;
            if(pageItems[0].text.length < 40){
                subTitle = pageItems.shift();
                subTitleSize = titleSize-5;
                subTitle.text = removeStartEndSpace(subTitle.text);
                subTitleFillInData = fillInText({...usingPreset.title}, subTitle.text, subTitleSize);
                subTitle.text = subTitleFillInData.text;
                titleFillInData.rect.y += (titleSize - (titleFillInData.rect.cy + subTitleFillInData.rect.cy + titleGap))/2
                subTitleFillInData.rect.y = titleFillInData.rect.y + titleFillInData.rect.cy + titleGap;
            }

            setPageNumStyle(slide, {
                fontFace: fontFace, 
                fontSize: 15, 
                fontBold: true, 
                fontUnderline: true, 
                textColor: bgiData.bColor, 
            });

            slide.backgroundColor(bgiData.bgc);
            addBgi(slide, bgiPath);
            nameList.push('!!background');

            slide.addShape({
                type: PPTX.ShapeTypes.RECTANGLE, 
                color: bgiData.bOrD == 'b' ? bgiData.dColor : bgiData.bColor, 
                ...usingPreset.shape
            });
            nameList.push('!!titleBox');

            slide.addText({
                value: title.text.toFullShape(), 
                fontFace: fontFace, 
                fontSize: titleSize, 
                fontBold: true, 
                textColor: bgiData.bOrD == 'b' ? bgiData.bColor : bgiData.dColor, 
                textWrap: 'none', 
                textVerticalAlign: 'center', 
                margin: 0, 
                ...usingPreset.title, 
                ...titleFillInData.rect
            });
            nameList.push('!!title');

            if(subTitleFillInData){
                slide.addText({
                    value: subTitle.text.toFullShape(), 
                    fontFace: fontFace, 
                    fontSize: subTitleSize, 
                    fontBold: true, 
                    textColor: bgiData.bOrD == 'b' ? bgiData.bColor : bgiData.dColor, 
                    textWrap: 'none', 
                    textVerticalAlign: 'center', 
                    margin: 0, 
                    shrinkText: true, 
                    ...usingPreset.title, 
                    ...subTitleFillInData.rect
                });
                nameList.push('!!subTitle');
            }

            let descriptions = [];
            pageItems.map((item, i) => {
                if(item.text !== undefined) descriptions.push(item.text);
                else{
                    console.log(item);
                    // console.log(item.embed);
                    if(item.type == 'embed'){
                        slide.addImage({
                            file: item.embed
                        });
                        nameList.push(`image_${i}`);
                    }
                }
            });
            descriptions = descriptions.map(row => fillInText(usingPreset.descriptions, row).text);
            let descriptionsCy = descriptions.join('\n').split('\n').length * fontSize;
            usingPreset.descriptions.y = usingPreset.descriptions.y + (fontSize - descriptionsCy)/2;
            slide.addText({
                value: descriptions.join('\n'), 
                fontFace: fontFace, 
                fontSize: fontSize, 
                textColor: bgiData.bOrD == 'b' ? bgiData.dColor : bgiData.bColor, 
                textWrap: 'none', 
                textVerticalAlign: 'center', 
                margin: 0, 
                ...usingPreset.descriptions, 
                cy: descriptionsCy
            });
            nameList.push(`!!descriptions`);

            resetElementCNvPr(slide, nameList);
        }
    }
    function defaultPage(pageItems){
        return slide => {
            let nameList = [];

            setPageNumStyle(slide, {
                fontFace: fontFace, 
                fontSize: 15, 
                fontBold: true, 
                fontUnderline: true, 
                textColor: bgiData.bColor, 
            });

            slide.backgroundColor(bgiData.bgc);
            addBgi(slide, bgiPath);
            nameList.push('!!background');

            pageItems.map((item, i) => {
                slide.addText({
                    value: item.text, 
                    x: gap, 
                    y: (fontSize+gap) * i, 
                    fontFace: fontFace, 
                    fontSize: fontSize, 
                    textColor: bgiData.bOrD == 'b' ? bgiData.dColor : bgiData.bColor, 
                    textWrap: 'none', 
                    textAlign: 'left', 
                    textVerticalAlign: 'center', 
                    // line: { color: '0000FF', dashType: 'dash', width: 1.0 }, 
                    margin: 0
                });
                nameList.push(`text_${i+1}`);
            });

            resetElementCNvPr(slide, nameList);
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
            if(['title', 'subtitle'].indexOf(item.type?.toLowerCase()) > -1) pagesData.push([]);
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
        Object.keys(pres.powerPointFactory.pptFactory.slideFactory.content)
        .filter(p => p.startsWith('ppt/slides/') && p.endsWith('.xml'))
        .forEach(p => {
            let $$ = [];
            let spTreeInner = pres.powerPointFactory.pptFactory.slideFactory.content[p]['p:sld']['p:cSld'][0]['p:spTree'][0];
            if(!settings.pageNum.value) spTreeInner['p:sp'].shift();
            else spTreeInner['p:sp'][0]['p:nvSpPr'][0]['p:cNvPr'][0].$.id = 1e3;
            for(let tagName in spTreeInner){
                spTreeInner[tagName].forEach(item => {
                    let id;
                    let subTagNameTable = {'p:sp': 'p:nvSpPr', 'p:pic': 'p:nvPicPr'};
                    if(Object.keys(subTagNameTable).indexOf(tagName) > -1){
                        id = item[subTagNameTable[tagName]][0]['p:cNvPr'][0].$.id;
                    }
                    $$.push({id, '#name': tagName, '#element': item});
                });
            }
            pres.powerPointFactory.pptFactory.slideFactory.content[p].temp$$ = $$
            .filter(n => n !== null)
            .sort((a, b) => parseInt(a.id !== undefined ? a.id : -1) - parseInt(b.id !== undefined ? b.id : -1));
        });
        pres.zipContent = zipContent.bind(pres);
    });
    return(pptx);
}
async function exportPptx(event, settings){
    let outputPath = `unpacked/output.pptx`;
    let pptx = await generatePptx(settings);
    await pptx.save(outputPath);
    return(outputPath);
}
function setPageNumStyle(slide, style){
    let pageNumElement = slide.powerPointFactory.pptFactory.slideFactory.content[`ppt/slides/${slide.name}.xml`]['p:sld']['p:cSld'][0]['p:spTree'][0]['p:sp'][0];
    let fld = pageNumElement['p:txBody'][0]['a:p'][0]['a:fld'][0];
    let styleData = {$: fld['a:rPr'][0].$};
    if(style.fontSize !== undefined) styleData.$.sz = style.fontSize * 100;
    if(style.fontBold !== undefined) styleData.$.b = style.fontBold ? 1 : 0;
    if(style.fontItalic !== undefined) styleData.$.i = style.fontItalic ? 1 : 0;
    if(style.fontUnderline !== undefined) styleData.$.u = style.fontUnderline ? 'sng' : 0;
    if(style.textColor) styleData['a:solidFill'] = [{'a:srgbClr': [{$: {val: style.textColor}}]}];
    if(style.fontFace){
        let fontFaceAttribute = [{$: {typeface: style.fontFace, pitchFamily: 0, charset: 0}}];
        styleData['a:latin'] = fontFaceAttribute;
        styleData['a:cs'] = fontFaceAttribute;
    }
    fld['a:rPr'][0] = styleData;
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