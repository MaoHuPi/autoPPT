# 新增主題文章並更新至subjectKeywords

## 步驟

1. 創建`subject`名稱的資料夾於`data/text/subjectArticle`下
2. 創建符合該`subject`的文章文字檔(utf-8)至該`subject`下
3. 執行`data/convertSubjectKeywords.js`
4. 複製`data/text/beforeTranslate.txt`的內容至Google翻譯
5. 清空`data/text/afterTranslate.txt`
6. 將翻譯結果(所有頁數)貼入`data/text/afterTranslate.txt`
3. 執行`data/translateKeywords.js`