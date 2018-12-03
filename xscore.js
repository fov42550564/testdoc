const Excel = require('exceljs');
const path = require('path');
const fs = require('fs-extra');
const readline = require('readline');
const program = require('commander');
const colors = require('colors');
const printf = require('printf');
const repeat = require('lodash.repeat');
const pinyin = require('./pinyin');

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
    async radiobox (title, disp, list, callback, width) {
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
            callback(list[index], index);
            break;
        }
    }
};
const dialog = new Dialog();

function isSamePinyin(py1, py2) {
    if (program.mohu) {
        for (let i = 0; i < py1.length; i++) {
            let ch1 = py1[i];
            let ch2 = py2[i];

                console.log("=====",i, ch1, ch2);
            if (ch1 === 'N' || ch1 === 'L') {
                if (ch2 !== 'N' && ch2 !== 'L') {
                    return false;
                }
            } else if (ch1 === 'F' || ch1 === 'H') {
                if (ch2 !== 'F' && ch2 !== 'H') {
                    return false;
                }
            } else {
                if (ch1 !== ch2) {
                    return false;
                }
            }
        }
        return true;
    }
    return py1 === py2;
}
function getFirstPinyinForWord(ch) {
    var uni = ch.charCodeAt(0);
    if (uni >= 48 && uni <= 57) {
        return "1";
    }
    if (uni >= 65 && uni <=90) {
        return uni;
    }
    if (uni >= 97 && uni <=122) {
        return String.fromCharCode(uni-32);
    }
    if (uni > 40869 || uni < 19968) {
        return "#";
    }
    return pinyin.charAt(uni - 19968);
}
function getFirstPinyin(str) {
    let ret = '';
    for (const ch of str) {
        ret += getFirstPinyinForWord(ch);
    }
    return ret;
}
async function setScoreForItem(filename, item, sc, wb, ws, map) {
    ws.getCell(`${item.colNumber}${item.rowNumber}`).value = sc;
    await wb.xlsx.writeFile(filename);
    console.log(`${item.name}: ${sc} 设置完成`);
    getInputValue(filename, wb, ws, map);
}
async function getInputValue(filename, wb, ws, map) {
    const str = await dialog.inputbox('请输入：'.INDEX_COLOR);
    if (str === 'q' || str === 'exit' || str === 'quit') {
        return;
    }
    const matches = str.match(/^([a-z]+)([0-9.]+)$/);
    if (!matches) {
        console.log('格式必须为：zs99');
        return getInputValue(filename, wb, ws, map);
    }
    const py = matches[1].toUpperCase();
    const sc = +matches[2];

    const list = map.filter(o=>isSamePinyin(py, o.py));

    if (list.length > 1) {
        dialog.radiobox("set", "选择一个来设置分数", list.map(o=>o.name), (unused, index)=>{
            setScoreForItem(filename, list[index], sc, wb, ws, map);
        });
    } else if (list.length === 1) {
        setScoreForItem(filename, list[0], sc, wb, ws, map);
    } else {
        console.log(`没有找到该学生`);
        getInputValue(filename, wb, ws, map);
    }
}
async function main(filename, min, max) {
    const map = [];
    const wb = new Excel.Workbook();
    await wb.xlsx.readFile(filename);
    const ws = wb.getWorksheet(1);
    const name1 = ws.getColumn('B'); //E
    const name2 = ws.getColumn('N'); //Q

    name1.eachCell((cell, rowNumber)=>{
        if (rowNumber >=min && rowNumber <= max && cell.value) {
            map.push({ name: cell.value, py: getFirstPinyin(cell.value), colNumber: 'E', rowNumber });
        }
    });
    name2.eachCell((cell, rowNumber)=>{
        if (rowNumber >=min && rowNumber <= max && cell.value) {
            map.push({ name: cell.value, py: getFirstPinyin(cell.value), colNumber: 'Q', rowNumber });
        }
    });
    getInputValue(filename, wb, ws, map);
}

program
.version('0.0.1')
.option('-b, --banji <1>', '设置班级，1：1班，2：2班，3：3班')
.option('-f, --filename <xx.xlsx>', '设置文件')
.option('-m, --mohu', '是否使用模糊音h=f/l=n')
.parse(process.argv);

if (!process.argv.slice(2).length) {
    program.help();
}

const { banji, filename } = program;
if (!filename) {
    console.log('必须设置xlsx文件');
    process.exit(0);
}

if (banji === '1') {
    main(filename, 4, 38);
} else if (banji === '2') {
    main(filename, 47, 81);
} else if (banji === '3') {
    main(filename, 90, 124);
}
