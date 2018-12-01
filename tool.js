const Excel = require('exceljs');
const path = require('path');
const fs = require('fs-extra');
const readline = require('readline');
const MAX_COL = 7;

function writeFile(mdname, line) {
    // console.log(mdname, line+'\n');
    fs.appendFileSync(mdname, line+'\n');
}
async function parseSingleExcelFile(filename) {
    const wb = await new Excel.Workbook();
    console.log(`开始分析 ${filename}.xlsx`);
    fs.emptyDirSync(path.join('docs', filename));
    await wb.xlsx.readFile(`excel/${filename}.xlsx`);
    await wb.xlsx.readFile(path.join('excel', `${filename}.xlsx`));
    let i = 0;
    wb.eachSheet(ws => {
        console.log(`开始写 ${ws.name}.md`);
        mdname = path.join('docs', filename, `${i++}_${ws.name}.md`);
        ws.eachRow((row, rowNumber)=>{
            if (rowNumber === 1) {
                writeFile(mdname, '| 编号 | 测试项 | 测试描述 | 测试结果 | 测试版本 | 测试时间 | 测试员 |');
                writeFile(mdname, '| :-: | :-: | :- | :- | :-: | :-: | :-: |');
            } else {
                let line = '|';
                row.eachCell((cell, colNumber)=>{
                    if (colNumber <= MAX_COL) {
                        let value = cell.value;
                        if (typeof value === 'string') {
                            value = value.replace(/\r\n|\r|\n/g, '<br>');
                        } else if (value.richText) {
                            value = value.richText.map(o=>o.text).join('');
                        }
                        line = `${line} ${value} |`;
                    }
                });
                writeFile(mdname, line);
            }
        });
    });
}
function parseAllExcelFile(root) {
    fs.readdirSync(root).forEach(file=>{
        parseSingleExcelFile(file.replace('.xlsx', ''));
    });
}
function readFileLine(filename, ws) {
    return new Promise(resolve=>{
        const rl = readline.createInterface({
            input: fs.createReadStream(filename)
        });
        let i = 0;
        rl.on('line', line=>{
            if (i++ > 1) {
                line = line.replace(/^\s*\|/, '').replace(/\|\s*$/, '');
                const row = line.split('|').map(o=>o.trim()).map(o=>o.replace(/<br>/g, '\r\n'));
                ws.addRow(row);
            }
        });
        rl.on('close', ()=>{
            resolve();
        });
    });
}
async function createExcelSheetFile(wb, name, filename) {
    const ws = wb.addWorksheet(name);
    ws.columns = [
        { header: '编号', width: 6, style: { alignment: { vertical: 'top', horizontal: 'center' } } },
        { header: '测试项', width: 20, style: { alignment: { vertical: 'top', horizontal: 'center' } } },
        { header: '测试描述', width: 50, style: { alignment: { wrapText: true } } },
        { header: '测试结果', width: 50, style: { alignment: { wrapText: true, vertical: 'top' } } },
        { header: '测试版本', width: 10, style: { alignment: { vertical: 'top', horizontal: 'center' } } },
        { header: '测试时间', width: 14, style: { alignment: { vertical: 'top', horizontal: 'center' } } },
        { header: '测试员', width: 10, style: { alignment: { vertical: 'top', horizontal: 'center' } } },
    ];
    await readFileLine(filename, ws);
}
async function createSingleExcelFile(name, dir, dist) {
    const wb = await new Excel.Workbook();
    const files = fs.readdirSync(dir);
    let cnt = 0;
    for (const file of files) {
        if (/\.md$/.test(file)) {
            cnt++;
            await createExcelSheetFile(wb, file.replace(/^\d+_|\.md$/g, ''), path.join(dir, file));
        }
    }
    cnt > 0 && await wb.xlsx.writeFile(path.join(dist, `${name}.xlsx`));
}
function createAllExcelFile(root, dist) {
    fs.emptyDirSync(dist);
    fs.readdirSync(root).forEach(file=>{
        const dir = path.join(root, file);
        fs.lstatSync(dir).isDirectory() && createSingleExcelFile(file, dir, dist);
    });
}
async function main() {
    parseAllExcelFile('excel');
    // createAllExcelFile('docs', 'excel');
}
main();
