const Excel = require('exceljs');
const path = require('path');
const fs = require('fs-extra');
const readline = require('readline');
const program = require('commander');
const colors = require('colors');
const printf = require('printf');
const moment = require('moment');
const repeat = require('lodash.repeat');

colors.setTheme({
    BORDER_COLOR: 'green',
    TITLE_COLOR: 'red',
    INDEX_COLOR: 'magenta',
    STRING_COLOR: 'blue',
    INPUTBOX_COLOR: 'blue',
    ERROR_COLOR: 'red',
    WARN_COLOR: 'yellow',
    DEBUG_COLOR: 'green',
});

class Dialog {
    constructor (simple) {
        this.HTAG = simple ? '-' : '\u2500';
        this.VTAG = simple ? '|' : '\u2502';
        this.TLTAG = simple ? '+' : '\u250c';
        this.TRTAG = simple ? '+' : '\u2510';
        this.MLTAG = simple ? '|' : '\u251c';
        this.MRTAG = simple ? '|' : '\u2524';
        this.BLTAG = simple ? '+' : '\u2514';
        this.BRTAG = simple ? '+' : '\u2518';

        this.DIALOG_WIDTH = 80;
        this.MSGBOX_WIDTH = 30;
    }
    error (args) {
        var ret = printf.apply(this, arguments);
        console.log(ret.ERROR_COLOR);
    }
    warn (args) {
        var ret = printf.apply(this, arguments);
        console.log(ret.WARN_COLOR);
    }
    debug (args) {
        var ret = printf.apply(this, arguments);
        console.log(ret.DEBUG_COLOR);
    }
    pause () {
        return new Promise(async (resolve) => {
            console.log('Press any key to continue...');
            process.stdin.setRawMode(true);
            process.stdin.resume();
            process.stdin.on('data', ()=>{
                resolve();
            });
        });
    }
    isAsc(code) {
        return code >= 0 && code <= 256;
    }
    isAscWord(code) {
        return (code >= 48 && code <= 57) || (code >= 65 && code <= 90) || (code >= 97 && code <= 122) || code === 95; //数字、字母、下划线
    }
    getRealLength(text) {
        let realLength = 0, len = text.length, charCode = -1;
        for (let i = 0; i < len; i++) {
            charCode = text.charCodeAt(i);
            if (this.isAsc(charCode)) {
                realLength += 1;
            } else {
                realLength += 2;
            }
        }
        return realLength;
    }
    getRealCutText(text, n) {
        let realLength = this.getRealLength(text);
        return text+repeat(' ', n-realLength);
    }
    cutLimitText (list, text, n) {
        let realLength = 0, len = text.length, preLen = -1, charCode = -1, needCut = false;
        for (let i = 0; i < len; i++) {
            charCode = text.charCodeAt(i);
            if (this.isAsc(charCode)) {
                realLength += 1;
            } else {
                realLength += 2;
            }
            if (preLen === -1 && realLength >= n) {
                preLen =  realLength === n ? i + 1 : i;
            } else if (realLength > n) {
                needCut = true;
                break;
            }
        }
        if (needCut) {
            let cutText = text.substr(0, preLen);
            let lastCode = cutText.charCodeAt(cutText.length-1);
            let nextCode = text.charCodeAt(preLen);

            if (this.isAscWord(lastCode) && this.isAscWord(nextCode)) {
                for (var j = 0; j < cutText.length; j++) {
                    if (!this.isAscWord(cutText.charCodeAt(cutText.length-1-j))) {
                        break;
                    }
                }
                if (j < cutText.length) {
                    preLen -= j;
                }
            }
            list.push(this.getRealCutText(text.substr(0, preLen), n));
            text = text.substr(preLen);
            if (text) {
                this.cutLimitText(list, text, n);
            }
        } else {
            list.push(this.getRealCutText(text, n));
        }
    }
    getBoder (width, layer = 'top', hasStart = true, hasEnd = true) {
        var headChar = (hasStart) ? (layer==='top' ? this.TLTAG : layer==='middle' ? this.MLTAG : this.BLTAG ) : this.HTAG;
        var tailChar = (hasEnd) ? (layer==='top' ? this.TRTAG : layer==='middle' ? this.MRTAG : this.BRTAG ) : this.HTAG;
        var str = headChar;
        for (var i=1; i<width; i++) {
            str += (i === width - 1) ? tailChar : this.HTAG;
        }
        return str;
    }
    showTopLineBoder (width) {
        width = width || this.DIALOG_WIDTH;
        console.log(this.getBoder(width).BORDER_COLOR);
    }
    showMiddleLineBoder (width) {
        width = width || this.DIALOG_WIDTH;
        console.log(this.getBoder(width, 'middle').BORDER_COLOR);
    }
    showBottomLineBoder (width) {
        width = width || this.DIALOG_WIDTH;
        console.log(this.getBoder(width, 'bottom').BORDER_COLOR);
    }
    showBoxTitle (title, width) {
        width = width || this.DIALOG_WIDTH;

        var len = width - this.getRealLength(title) - 2;
        var flen = parseInt(len / 2);
        var flen1 = flen + (len & 1);

        console.log(this.getBoder(flen, 'top', true, false).BORDER_COLOR +' '+ title.TITLE_COLOR +' '+ this.getBoder(flen1, 'top', false, true).BORDER_COLOR);
    }
    showBoxDiscription (disp, width) {
        width = width || this.DIALOG_WIDTH;
        this.showBoxString({head: '[Usage]:', text: disp}, width);
    }
    showBoxString (item, width) {
        var {head = '', text = ''} = item;
        text = ' ' + text;
        width = width || this.DIALOG_WIDTH;
        var in_width = width - 3;
        var list = [];
        this.cutLimitText(list, head+text, in_width);
        for (var i in list) {
            if (i == 0) {
                console.log(this.VTAG.BORDER_COLOR + ' ' + head.INDEX_COLOR + list[i].replace(head, '').STRING_COLOR + this.VTAG.BORDER_COLOR);
            } else {
                console.log(this.VTAG.BORDER_COLOR + ' ' + list[i].STRING_COLOR + this.VTAG.BORDER_COLOR);
            }
        }
    }
    msgbox (title, list, width) {
        width = width || this.DIALOG_WIDTH;
        var len = list.length;
        this.showBoxTitle(title, width);
        for (var i = 0; i < len; i++) {
            this.showBoxString(list[i], width);
        }
        this.showBottomLineBoder(width);
    }
    inputbox (info, defaultValue = '') {
        return new Promise(async (resolve) => {
            info = info.INPUTBOX_COLOR;
            if (defaultValue) {
                info += ('[default:' + defaultValue + ']:').INDEX_COLOR;
            }
            var rl = readline.createInterface({
                input: process.stdin,
                output: process.stdout,
            });
            rl.question(info, (ret) => {
                ret = ret.trim() || defaultValue;
                rl.close();
                resolve(ret);
            });
        });
    }
    listbox (title, disp, list, width) {
        width = width || this.DIALOG_WIDTH;
        var len = list.length;
        this.showBoxTitle(title, width);
        this.showBoxDiscription(disp, width);
        this.showMiddleLineBoder(width);
        for (var i = 0; i < len; i++) {
            this.showBoxString({head: i + ':', text: list[i]}, width);
        }
        this.showBottomLineBoder(width);
    }
    async radiobox (title, disp, list, callback, isUpdate, width) {
        width = width || this.DIALOG_WIDTH;
        var len = list.length;
        while (true) {
            if (len == 0) {
                this.msgbox('complete', [{text: 'have done all'}], this.MSGBOX_WIDTH);
                return;
            }
            this.listbox(title, disp, list, width);
            var ret = await this.inputbox('please input need ' + title.TITLE_COLOR + ' index,' + 'exit(q)'.INDEX_COLOR);
            var index = ret.trim();
            if (!index.length) {
                this.error('null is not allowed');
                continue;
            }
            if (index == 'q') {
                break;
            }
            if (isNaN(index)) {
                this.error('must input a number');
                continue;
            }
            if (index < 0 || index >= len) {
                this.error('the select number is out of range');
                continue;
            }
            callback(list[index]);
            if (isUpdate) {
                list.splice(index, 1);
                len--;
            }
            await this.pause();
        }
    }
};

const dialog = new Dialog();
const MAX_COL = 7;

function writeFile(mdname, line) {
    // console.log(mdname, line+'\n');
    fs.appendFileSync(mdname, line+'\n');
}
async function parseSingleExcelFile(filename, config) {
    const wb = new Excel.Workbook();
    console.log(`开始分析 ${filename}.xlsx`);
    fs.emptyDirSync(path.join('docs', filename));
    await wb.xlsx.readFile(path.join('excel', `${filename}.xlsx`));
    let i = 0;
    const pages = [];
    wb.eachSheet(ws => {
        console.log(`开始写 ${ws.name}.md`);
        config && pages.push({ name: ws.name, path: `${filename}/${i}_${ws.name}.md` });
        mdname = path.join('docs', filename, `${i++}_${ws.name}.md`);
        ws.eachRow((row, rowNumber)=>{
            if (rowNumber === 1) {
                writeFile(mdname, '| 编号 | 测试项 | 测试描述 | 测试结果 | 测试版本 | 测试时间 | 测试员 |');
                writeFile(mdname, '| :-: | :-: | :- | :- | :-: | :-: | :-: |');
            } else {
                let line = '|';
                for (let i=1; i<=MAX_COL; i++) {
                    let value = row.getCell(i).value;
                    if (value != null) {
                        if (value.richText) {
                            value = value.richText.map(o=>o.text).join('');
                        }
                        if (typeof value === 'string') {
                            value = value.replace(/\r\n|\r|\n/g, '<br>');
                        } else if (value instanceof Date) {
                            value = moment(value).format('YYYY-MM-DD');
                        }
                    } else {
                        value = '';
                    }
                    line = `${line} ${value} |`;
                }
                writeFile(mdname, line);
            }
        });
    });
    config && config.menus.push({ name: filename, groups: [ { pages } ] });
}
async function parseAllExcelFile(root, config) {
    const files = fs.readdirSync(root);
    config && (config.menus = []);
    for (const file of files) {
        if (/^[^~.\.].*\.xlsx$/.test(file)) {
            await parseSingleExcelFile(file.replace('.xlsx', ''), config);
        }
    }
    if (config) {
        fs.writeFileSync('./document/config.js', `module.exports = ${JSON.stringify(config, null, 4)}`);
        let files = fs.readdirSync('./document/docs');
        for (const file of files) {
            const _file = path.join('./document/docs', file);
            if (fs.lstatSync(_file).isSymbolicLink()) {
                fs.removeSync(_file);
            }
        }
        files = fs.readdirSync('./docs');
        for (const file of files) {
            const _file = path.join('./docs', file);
            if (fs.lstatSync(_file).isDirectory()) {
                fs.ensureSymlinkSync(_file, path.join('./document/docs', file));
            }
        };
    }
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
    if (cnt > 0) {
        await wb.xlsx.writeFile(path.join(dist, `${name}.xlsx`));
        console.log(`生成 ${name}.xlsx 成功`);
    }
}
async function createAllExcelFile(root, dist) {
    fs.emptyDirSync(dist);
    const files = fs.readdirSync(root);
    for (const file of files) {
        const dir = path.join(root, file);
        if (fs.lstatSync(dir).isDirectory()) {
            await createSingleExcelFile(file, dir, dist);
        }
    }
}

program
.version('0.0.1')
.option('-e, --excel', '使用docs中的一个文件夹名生产excel文件在excel目录中，如果为all则生成所有的excel文件')
.option('-d, --docs', '使用excel中的一个excel文件生成md文件在mdoc中，如果为all这生成所有的md文件')
.option('-c, --config', '生产config文件，只有 -d all 的时候有效【测试人员禁用】')
.parse(process.argv);

if (!process.argv.slice(2).length) {
    program.help();
}

const {  excel, docs,config } = program;

if (docs) {
    const files = fs.readdirSync('excel');
    const list = ['all'];
    for (const file of files) {
        if (/^[^~.\.].*\.xlsx$/.test(file)) {
            list.push(file);
        }
    }
    dialog.radiobox("change", "选择一个excel文件来转化为md文件在mdoc中", list, async (file)=>{
        if (file === 'all') {
            if (config) {
                await parseAllExcelFile('excel', require('./document/config.js'));
            } else {
                await parseAllExcelFile('excel');
            }
        } else {
            await parseSingleExcelFile(file.replace('.xlsx', ''));
        }
        process.exit(0);
    });
} else if (excel) {
    const files = fs.readdirSync('docs');
    const list = ['all'];
    for (const file of files) {
        const dir = path.join('docs', file);
        if (fs.lstatSync(dir).isDirectory()) {
            list.push(file);
        }
    }
    dialog.radiobox("change", "选择一个docs中的文件夹来转化为excel文件在excel中", list, async (file)=>{
        if (file === 'all') {
            await createAllExcelFile('docs', 'excel');
        } else {
            await createSingleExcelFile(file, path.join('docs', file), 'excel');
        }
        process.exit(0);
    });
}
