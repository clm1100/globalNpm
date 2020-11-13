const XLSX = require('xlsx')
const XLSXStyle = require('xlsx-style')
const fs = require('fs');
function deal(p){
    const arr = fs.readdirSync(p);
    const newarr = arr.filter(e => {
        return !e.includes('清单') && e.includes('pdf')
    })

    let tableData = newarr.map(e => {
        let list = e.split('-');
        let l = list.length;
        let code = null
        let pingzheng = null;
        let name = null
        let date = null
        let money = null;
        if (l == 4) {
            code = list[1];
            name = list[3].split('.')[0];
            money = list[2];
        } else if (l == 3) {
            pingzheng = list[0];
            name = list[2].split('.')[0];
            money = list[1]
        }
        return {
            code, pingzheng, money, date, name
        }
    })

    tableData.sort((a, b) => a.name.localeCompare(b.name));

    let table = tableData.map((e, i) => {
        return [
            i + 1,
            e.code,
            e.pingzheng,
            e.name,
            e.date,
            e.money
        ]
    })



    let arrHeader = ['序号', '发票号码', '凭证号', '报销人', '报销日期', '发票金额'];


    table.unshift(arrHeader)
    table.unshift([{
        v: '北京华亿创新信息技术有限公司电子发票报销台账（2020年）',
        s: {
            font: {
                name: '宋体',
                sz: 24,
                bold: true,
                color: { rgb: "FFFFAA00" }
            },
            alignment: { horizontal: "center", vertical: "center", wrap_text: true },
            fill: { bgcolor: { rgb: 'ffff00' } }
        }
    }, null, null, null, null, null]);


    //1、定义导出文件名称
    var filename = "write.xlsx";
    // 定义导出数据
    var data = table
    // 定义excel文档的名称
    var ws_name = "2020年电子发票台账";
    // 初始化一个excel文件
    var wb = XLSX.utils.book_new();
    // 初始化一个excel文档，此时需要传入数据
    var ws = XLSX.utils.aoa_to_sheet(data);

    ws['!merges'] = [
        // 设置A1-C1的单元格合并
        { s: { r: 0, c: 0 }, e: { r: 0, c: 5 } }
    ];
    // 将文档插入文件并定义名称
    XLSX.utils.book_append_sheet(wb, ws, ws_name);
    // 执行下载
    // XLSX.writeFile(wb, filename);
    XLSXStyle.writeFile(wb, filename);
}
module.exports = deal;