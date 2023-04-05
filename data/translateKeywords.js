// basic import
const fs = require('fs');
const path = require('path');
String.prototype.replaceAll = function (target, payload) {
    let regex = new RegExp(target, 'g');
    return(this.valueOf().replace(regex, payload));
};

let countData;
try{
    countData = fs.readFileSync(path.join(__dirname, 'json/subjectKeywords.json'), 'utf-8');
    countData = JSON.parse(countData);
}
catch(err){countData = {};}

function getTextData(name){
    let data = [];
    try{
        data = fs.readFileSync(path.join(__dirname, `text/${name}.txt`), 'utf-8');
        data = data.replaceAll(/[\r\n]/g, '').replaceAll('{', '').split('}');
    }
    catch(err){
        data = [];
        console.log(`File Not Found! (${name})`);
    }
    return(data);
}
let beforeTranslate = getTextData('beforeTranslate');
let afterTranslate = getTextData('afterTranslate');

beforeTranslate.map((keyword, i) => {
    if(countData[keyword] !== undefined){
        let translatedList = countData[keyword][0];
        let newTranslate = afterTranslate[i];
        if(newTranslate !== keyword && translatedList.indexOf(newTranslate) === -1 && newTranslate !== undefined && newTranslate !== null){
            translatedList.push(newTranslate);
        }
    }
});

fs.writeFileSync(path.join(__dirname, 'json/subjectKeywords.json'), JSON.stringify(countData, true, 4), 'utf-8');
// fs.writeFileSync(path.join(__dirname, 'text/beforeTranslate.txt'), '', 'utf-8');
// fs.writeFileSync(path.join(__dirname, 'text/afterTranslate.txt'), '', 'utf-8');