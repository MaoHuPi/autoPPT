(() => {
    const html = $('html');
    const pages = $('#pages');
    let pageNow = 0;
    function pageChange(rate){
        if(typeof rate != typeof 0) rate = this.rate; 
        if(pageNow + rate >= 0 && pageNow + rate < pages.childElementCount) pageNow += rate;
        html.style.setProperty('--pageNow', pageNow);
    }
    $('#floatButton-back').addEventListener('click',  pageChange.bind({rate: -1}));
    $('#floatButton-next').addEventListener('click',  pageChange.bind({rate: 1}));
})();