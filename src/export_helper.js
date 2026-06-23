(function() {
    'use strict';

    function isInPyWebView() {
        return !!(window.pywebview && window.pywebview.api);
    }

    function fallbackDownloadText(text, filename, fileType) {
        var blob = null;
        if (fileType === 'json') {
            blob = new Blob([text], { type: 'application/json;charset=utf-8' });
        } else {
            blob = new Blob(["\uFEFF" + text], { type: 'text/plain;charset=utf-8' });
        }
        var url = URL.createObjectURL(blob);
        var a = document.createElement('a');
        a.href = url;
        a.download = filename + '.' + fileType;
        document.body.appendChild(a);
        a.click();
        document.body.removeChild(a);
        URL.revokeObjectURL(url);
    }

    function fallbackDownloadWorkbook(workbook, filename, fileType) {
        if (typeof XLSX !== 'undefined' && XLSX.writeFile) {
            XLSX.writeFile(workbook, filename + '.' + fileType);
        }
    }

    window.exportWithSaveDialog = async function(workbook, filename, fileType) {
        fileType = fileType || 'xlsx';
        try {
            if (isInPyWebView()) {
                if (typeof XLSX === 'undefined') {
                    throw new Error('XLSX 库未加载');
                }
                var wbout = XLSX.write(workbook, { bookType: fileType, type: 'base64' });
                var result = await window.pywebview.api.save_base64_file(wbout, filename, fileType);
                if (result && result.success) {
                    alert('文件已保存至：\n' + result.path);
                    return true;
                } else if (result && result.cancelled) {
                    return false;
                } else {
                    throw new Error((result && result.error) || '保存失败');
                }
            } else {
                fallbackDownloadWorkbook(workbook, filename, fileType);
                return true;
            }
        } catch (e) {
            console.error('导出失败:', e);
            alert('导出失败：' + (e && e.message ? e.message : e));
            return false;
        }
    };

    window.saveTextWithDialog = async function(text, filename, fileType) {
        fileType = fileType || 'txt';
        try {
            if (isInPyWebView()) {
                var result = await window.pywebview.api.save_file(text, filename, fileType);
                if (result && result.success) {
                    alert('文件已保存至：\n' + result.path);
                    return true;
                } else if (result && result.cancelled) {
                    return false;
                } else {
                    throw new Error((result && result.error) || '保存失败');
                }
            } else {
                fallbackDownloadText(text, filename, fileType);
                return true;
            }
        } catch (e) {
            console.error('保存失败:', e);
            alert('保存失败：' + (e && e.message ? e.message : e));
            return false;
        }
    };

    window.saveBase64WithDialog = async function(base64Data, filename, fileType) {
        try {
            if (isInPyWebView()) {
                var result = await window.pywebview.api.save_base64_file(base64Data, filename, fileType);
                if (result && result.success) {
                    alert('文件已保存至：\n' + result.path);
                    return true;
                } else if (result && result.cancelled) {
                    return false;
                } else {
                    throw new Error((result && result.error) || '保存失败');
                }
            } else {
                var bin = atob(base64Data);
                var bytes = new Uint8Array(bin.length);
                for (var i = 0; i < bin.length; i++) bytes[i] = bin.charCodeAt(i);
                var blob = new Blob([bytes]);
                var url = URL.createObjectURL(blob);
                var a = document.createElement('a');
                a.href = url;
                a.download = filename + '.' + fileType;
                document.body.appendChild(a);
                a.click();
                document.body.removeChild(a);
                URL.revokeObjectURL(url);
                return true;
            }
        } catch (e) {
            console.error('保存失败:', e);
            alert('保存失败：' + (e && e.message ? e.message : e));
            return false;
        }
    };
})();