

const fs = require('fs').promises,
    path = require('path'),
    PDFParser = require("pdf2json");

const ExcelJS = require('exceljs');

(async function () {
    // console.log(process.argv)
    const [node, script, dir] = process.argv;
    const dirPath = path.join(process.cwd(), dir)
    // console.log(dirPath)
    const files = await fs.readdir(dirPath)
    // console.log(files)



    const workbook = new ExcelJS.Workbook();
    const worksheet = workbook.addWorksheet('My Sheet');
    // worksheet.addRow()
    worksheet.columns = [
        {header:'编号',key:'index'},
        {header:'项目名称',key:'items'},
        {header:'发票代码',key:'code'},
        {header:'发票号码',key:'number'},
        {header:'发票日期',key:'date'},
        {header:'发票金额',key:'amount'},
    ]

    let pages = []
    for (let i = 0, len = files.length; i < len; i++) {
        let pdfFile = path.join(dirPath, files[i])
        let parseJson = await readPDF(pdfFile)
        let jsonStr = JSON.stringify(parseJson);
        // fs.writeFile('temp.json', jsonStr)
        let textBoxes = getTextBoxes(parseJson)
        // console.log(textBoxes)
        // fs.writeFile('result.json', JSON.stringify(textBoxes))
        // fs.writeFile('temp.html',makeHTML(textBoxes))
        let obj = getPageData(textBoxes)
        obj.index = i+1
        pages.push(obj)
        // console.log(obj)
        // break;
        worksheet.addRow(obj);
    }


    await workbook.xlsx.writeFile("result.xlsx");
})().catch(e => console.error(e)).finally(() => {
    console.log('done!')
})

const BillType = {
    '广东增值税电子普通发票':{
        name: '广东增值税电子普通发票',
        code: [Scale(908, 53), Scale(980, 67)],
        number: [Scale(905, 75), Scale(1130, 97)],
        date: [Scale(902, 110), Scale(1038, 134)],
        amount: [Scale(867, 522), Scale(1096,550)],
        items: [Scale(49, 304), Scale(342, 481)],
    }
};

function getBillType(textBoxes){
    // console.log(textBoxes)
    let result= textBoxes.find(item => /普通发票/.test(item.t))
    return result ? result.t : ''
}

function getPageData(textBoxes) {
    let obj = {index:"", code:'',number:'',date:'',amount:'',items:''};
    let type = getBillType(textBoxes)
    let points = BillType[type]
    if(!points){
        return obj
    }

    for(let k in obj){
        if(points[k]){
            obj[k]=getTextFromRectangle(textBoxes,...points[k])
        }
    }

    format(obj)
    return obj
}

function format(obj) {
    for (let k in obj) {
        if (obj[k].length == 0) {
            obj[k] = ''
            continue;
        }
        if (k == 'amount') {
            obj[k] = obj[k][0].replace(/[￥¥]/, '').trim()
        } else if (k == 'items') {
            obj[k] = obj[k].join('\n')
        } else if(Array.isArray(obj[k])) {
            obj[k] = obj[k][0]
        }
    }
}

function makeHTML(textBoxes, scale) {
    let str = ''
    scale = scale > 0 ? scale : 30
    textBoxes.forEach(item => {
        str += `<div style="left:${item.x * scale}px;top:${item.y * scale}px">${item.t}</div>`
    })
    return `<!DOCTYPE html>
    <style>
    * {
        margin:0;
        padding:0;
    }
    .page{
        position: relative;
        width:1300px;
        height:825px;
        background-color:#aaa;
        font-size:12px;
    }
    div{
        position:absolute;
        width: fit-content;
    }</style>
    <body>
        <div class="page">${str}</div>
    </body>`
}

function Scale(x, y) {
    // 830,530
    // 648,43 804,62

    // 1300,827
    // 1013,63 1244,95
    // return {x:x*38.974/1300,y:y*24.75/825}
    return { x: x / 30, y: y / 30 }
}

function getTextFromRectangle(arr, lt, rb) {
    // console.log(lt,rb)
    let textBoxes = []
    arr.forEach(item => {
        if (item.x > lt.x && item.y > lt.y && item.x < rb.x && item.y < rb.y) {
            // console.log(item)
            textBoxes.push(item)
        }
    })
    return textBoxes.sort((a, b) => a.x - b.x).map(e => e.t)
}

function getTextBoxes(data) {
    if (!data.Pages) {
        return
    }
    let textBoxes = [];
    data.Pages.forEach(page => {
        if (!page.Texts) {
            return
        }
        page.Texts.forEach(text => {
            if (!text.R) {
                return
            }
            let obj = {
                x: text.x,
                y: text.y,
                t: '',
            }
            text.R.forEach(r => {
                if (r.T) {
                    obj.t += decodeURIComponent(r.T)
                }
            })
            textBoxes.push(obj)
        })
    })
    return textBoxes
}

function readPDF(filepath) {
    const pdfParser = new PDFParser();
    return new Promise((resolve, reject) => {
        // pdfParser.on("pdfParser_dataError", errData => console.error(errData.parserError) );
        pdfParser.on("pdfParser_dataError", errData => reject(errData));
        pdfParser.on("pdfParser_dataReady", pdfData => {
            resolve(pdfData)
        });
        pdfParser.loadPDF(filepath);
    });
}