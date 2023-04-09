## 前置作業

### docx檔案結構

副檔名改`.zip`後解壓縮的結果如下

```js
/*
 * {} 任意名稱
 * [] 可能不存在
 * () 細部說明
 * `` 說明內的檔案會資料夾名稱
 * "" 內部文字部適用上述規則
 */

{unpack}.zip/
    _rels/
        .rels => 完全相同(指定xml版本及規格、內容的目標檔案)
    word/
        _rels/
            document.xml.rels => 大量不同(指定`document.xml`與`media/`及其他檔案的對應關係)
            [fontTable.xml.rels] => 大量不同(指定`fontTable.xml`與`fonts/`的對應關係)
        media/
            {image1.jpg} => 個別不同(內文穿插的非文字之檔案)
        [fonts]/
            {Arimo-bold.ttf} => 個別不同(內文套用到的字體檔案)
        theme/
            theme1.xml => 完全相同(描述字體及顏色的相關設定)
        document.xml => 大量不同(主要內容存放處，包含文字、格式等)
        fontTable.xml => 部分不同(描述了各使用到的字體的相關參數)
        numbering.xml => 部分不同(應該是在描述各文字樣式preset的相關參數)
        settings.xml => 完全相同(對照用設定)
        styles.xml => 略為不同(描述了各文字樣式preset的實際參數)
    "[Content_Types].xml" => 略為不同(有無使用ttf ? 有ttf規格說明標籤 : 反之)
```

### 初步內容提取

```js
contentData = [...document.body.querySelectorAll('p')]
    .map(p => {return({
        style: p.querySelector('pPr pStyle')?.getAttribute('w:val'), 
        text: [...p.querySelectorAll('t')].map(t => t.innerHTML), 
    });})
    .filter(data => data.text.length >= 1);
contentList = contentData.map(item => item.text);
console.log(contentList);
copy(contentList.join('\x01').replaceAll(/[\t\n]/g, '').replaceAll('\x01', '\n'));
```

## 功能設計

### 自訂調整參數構思

報告用
 => 進度條
歷程用
 => 標頁碼

## TODO

- [x] p retype
- [x] test and debug titlePage posData
- [x] fillInImage and fillInText function
- [x] titlePage posData => combine
- [x] titlePage posData add subtitle and description(p)
- [ ] defaultPage add image and table
- [ ] mask and maskColor
- [ ] image absPos function (fillMode can be overflow, fillAll or inside)
- [ ] image(not in table) can be front or back of the mask
- [ ] elements rotate (rotatedBox match back to oriBox)