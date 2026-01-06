/* static/js/taskpane.js v4.7 - 智能表格全选吸取 */

Office.onReady((info) => {
    if (info.host === Office.HostType.Word||
    info.host === Office.HostType.PowerPoint ||
    info.host === Office.HostType.Excel) {
            import('./ui.js').then(({ initializeTaskpane }) => {
                initializeTaskpane();
            });
    }
});
