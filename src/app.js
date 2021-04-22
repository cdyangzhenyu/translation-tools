window.onload = function () {

    var $ = function (select, ctx) {
        ctx = ctx || document;
        return ctx.querySelector(select);
    };

    var file = $('#file');
    var drop = $('#drop');
    var edit = $('#edit');
    var view = $('#view');
    var derive = $('#derive');

    var excelView = $('#excel-view');
    function fanyi(q) {
        
        return data;
    }

    /**
     * 读取
     */
    function readExcelFile(filedata, cb) {
        // https://developer.mozilla.org/zh-CN/docs/Web/API/FileReader
        var reader = new FileReader();

        var types = [
            'application/vnd.ms-excel',
            'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
        ];

        if (types.indexOf(filedata.type) === -1) {
            return alert('文件类型不是 Excel 格式');
        }

        reader.readAsBinaryString(filedata);
        reader.onload = function (e) {
            // 解析数据
            var bstr = e.target.result;
            var wb = XLSX.read(bstr, { type: 'binary', cellStyles: true });
            console.log(wb);
            
            
            //console.log(data);
            // 只取第一个 sheet
            // var wsname = wb.SheetNames[0];
            // var ws = wb.Sheets[wsname];
            // 渲染
            typeof cb === 'function' && cb(wb);
        };
    }

    /**
     * 渲染数据
     */
    function render(element, filedata) {
        readExcelFile(filedata, function (wb) {
            // 只取第一个 sheet
            var wsname = wb.SheetNames[0];
            var ws = wb.Sheets[wsname];
            var words = wb.Strings;
            console.log(words.length);
            //调用百度翻译接口
            var appid = 'your appid';
            var key = 'your key';
            var salt = (new Date).getTime();
            // 多个query可以用\n连接  如 query='apple\norange\nbanana\npear'
            var from = 'en';
            var to = 'zh';
            var jq=jQuery.noConflict();
            var results = new Array();
            for (var i = 0, len = words.length; i < len; i++) {
                var query = words[i].t;
                var str1 = appid + query + salt +key;
                var sign = MD5(str1);
                var new_dict = {};
                jq.ajax({
                    url: 'http://api.fanyi.baidu.com/api/trans/vip/translate',
                    type: 'get',
                    dataType: 'jsonp',
                    data: {
                        q: query,
                        appid: appid,
                        salt: salt,
                        from: from,
                        to: to,
                        sign: sign,
                        dict: 0
                    },
                    success: function (data) {
                        console.log(data); 
                        console.log(data.trans_result);
                        //console.log(JSON.parse(data.trans_result[0].dict));
                        var dict = data.trans_result[0].dict;
                        var word = data.trans_result[0].src;
                        var means = data.trans_result[0].dst;
                        var ph_am = '';
                        if(dict != ''){
                            means = '';
                            dict = JSON.parse(dict);
                            var symbols = dict.word_result.simple_means.symbols[0];
                            symbols.parts.forEach(function (item) {
                                means += item.part + ' ' + item.means.toString() + '\n';  
                            });
                            //console.log(means);
                            ph_am = '['+symbols.ph_am+']';
                        }
                        new_dict = {
                            'Word': word,
                            'Pronounciation': ph_am,
                            'Definition': means
                        }
                        results.push(new_dict);
                        console.log(results);
                        if(results.length == words.length) {
                            ws = XLSX.utils.json_to_sheet(results);
                            // 渲染
                            element.innerHTML = XLSX.utils.sheet_to_html(ws);
                        }
                    } 
                });
            }
            
        })
    }

    /**
     * 上传的文件
     */
    file.onchange = function (event) {
        var files = event.target.files;

        if (files && files[0]) {
            render(excelView, files[0]);
        }
    }


    // 把 string 转为 ArrayBuffer
    function s2ab(str) {
        var buf = new ArrayBuffer(str.length);
        var _view = new Uint8Array(buf);
        for (var i = 0, len = str.length; i < len; i++) {
            _view[i] = str.charCodeAt(i) & 0xFF;
        }
        return buf;
    }

    // 根据表格内容，生成 Excel 文件
    derive.onclick = function (event) {
        var table = $('table', view);
        var sheet = XLSX.utils.table_to_sheet(table);

        sheet['A1'] = Object.assign(sheet['A1'], {
            // 样式？
            s: {
                fill: {
                    fgColor: { rgb: "FFFF0000" }
                }
            },
        });

        var wb = XLSX.utils.book_new({ cellStyles: true });
        XLSX.utils.book_append_sheet(wb, sheet, "SheetJS");
        // 渲染
        var wbout = XLSX.write(wb, { type: "binary", bookType: "xlsx" });
        // 保存 - https://github.com/eligrey/FileSaver.js
        saveAs(new Blob([s2ab(wbout)], { type: "application/octet-stream" }), Date.now() + ".xlsx");
    }
}

