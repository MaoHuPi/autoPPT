// basic import
const fs = require('fs');
const path = require('path');
String.prototype.replaceAll = function (target, payload) {
    let regex = new RegExp(target, 'g');
    return(this.valueOf().replace(regex, payload));
};

let ignoreWords;
ignoreWords = fs.readFileSync(path.join(__dirname, 'text/ignoreWords.txt'), 'utf-8');
ignoreWords = ignoreWords.replaceAll('\n', '|');
ignoreWords = new RegExp(` (${ignoreWords}) `, 'g');

function removeIgnore(content){
    content = ` ${content} `
    .toLowerCase()
    .replaceAll(/([\!-\/]|[0-9]|[\:-\@]|[\[-\`])/g, ' ')
    .replaceAll(/[.,?!:;'"‘’“”,]/g, ' ')
    .replaceAll(/[\r\n]/g, ' ')
    .replaceAll(/ +/g, ' ')
    .replaceAll(' ', '  ')
    .replaceAll(ignoreWords, '  ');
    return(content);
}

let oldWordCount = {};
try{
    oldWordCount = fs.readFileSync(path.join(__dirname, 'json/subjectKeywords.json'), 'utf-8');
    oldWordCount = JSON.parse(oldWordCount);
}
catch(err){
    oldWordCount = {};
    console.log('File Not Found! (subjectKeywords)');
}

let articlePath = path.join(__dirname, 'text/subjectArticle/');
let countData = {};
let wordCount = {};
if(fs.existsSync(articlePath)){
    for(let subject of fs.readdirSync(articlePath) || []){
        let subjectPath = path.join(articlePath, subject);
        wordCount[subject] = 0;
        for(let file of fs.readdirSync(subjectPath) || []){
            let content = fs.readFileSync(path.join(subjectPath, file), 'utf-8');
            content = String(content);
            content = content.replaceAll(/.*article source.[^\n]*\n+/g, '');
            content = removeIgnore(content);
            content = content.split(' ');
            content.forEach(word => {
                if(countData[word] === undefined) countData[word] = [[], {}];
                if(countData[word][1][subject] === undefined) countData[word][1][subject] = 0;
                countData[word][1][subject] += 1;
            });
            wordCount[subject] += content.length;
        }
    }
}
let toBeTranslated = [];
for(let keyword in countData){
    let subjectValues = countData[keyword][1];
    for(let subject in subjectValues){
        subjectValues[subject] = subjectValues[subject]/wordCount[subject]*100;
    }
    let totalCount = Object.entries(subjectValues).reduce((t, o) => t + o[1], 0);
    for(let subject in subjectValues){
        subjectValues[subject] /= totalCount;
    }
    if(countData[keyword][0].length == 0) toBeTranslated.push(keyword);
}
for(let keyword in countData){
    if(oldWordCount[keyword] !== undefined){
        let n = countData[keyword][1];
        let o = oldWordCount[keyword][1];
        for(let subject of new Set([...Object.keys(n), ...Object.keys(o)])){
            n[subject] = (n[subject] !== undefined ? n[subject] : 0);
            o[subject] = (o[subject] !== undefined ? o[subject] : 0);
            n[subject] = (n[subject] + o[subject]) / 2;
        }
    }
}
fs.writeFileSync(path.join(__dirname, 'json/subjectKeywords.json'), JSON.stringify(countData, true, 4), 'utf-8');

let toBeTranslated_text = `{${toBeTranslated.join('}\n{')}}`;
fs.writeFileSync(path.join(__dirname, 'text/beforeTranslate.txt'), toBeTranslated_text, 'utf-8');