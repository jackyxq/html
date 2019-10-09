(function () {
    "use strict";

    // 每次加载新页面时都必须运行初始化函数
    Office.initialize = function (reason) {
        $(document).ready(function () {
            $('.btn-control').click(previewHtml);

            var url = Office.context.document.settings.get("hh");
            loadUrl(url);
            $('.webpage-frame').load(function () {
                var f = $('.webpage-frame');
                saveUrl(f[0].src);
            });
        });
    };

    // 从当前选择的文档内容中读取数据并显示通知
    function previewHtml() {
        var webpage = $('.webpage-decorator');
        var ishow = !webpage.is(":hidden");
        if (ishow) {
            webpage.hide();
            $('.btn-control span').text("预览");
        }
        else {
            var url = $('.form-control').val();
            loadUrl("https://" + url);
        }
    }

    function loadUrl(url) {
        if (url == null || url.length <= 8) return; /* https:// 的长度是8 */

        $('.webpage-decorator').show();
        $('.webpage-frame').attr("src", url);
        $('.btn-control span').text("编辑");
    }

    function saveUrl(url) {
        if (url.length <= 8) return;
        Office.context.document.settings.set("hh", url);
        Office.context.document.settings.saveAsync(function (r) {
            console.log(r);
        });
    }
    
})();