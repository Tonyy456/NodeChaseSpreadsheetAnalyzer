const ExcelJS = require('exceljs');
const dayjs = require('dayjs');
const fs = require('fs');
var customParseFormat = require('dayjs/plugin/customParseFormat')
dayjs.extend(customParseFormat)

const main = async () => {
    console.log("##### Running analysis on input files \n")
    const filePath = process.argv[2]; 
    const outputExists = fs.existsSync(process.argv[3]);

    /* CREATE WORKBOOK */
    const workbook = new ExcelJS.Workbook();
    if(outputExists) {
        await workbook.xlsx.readFile(process.argv[3]);
        console.log(`##### Reading XLSX file: ${process.argv[3]}`)
    } 
    const worksheet = await workbook.csv.readFile(filePath);
    worksheet.name = `${dayjs().format('YY-MM-DD')} Analysis @${dayjs().format('hh_mm_ss')}`
    workbook.calcProperties.fullCalcOnLoad = true;

    /* Constants */
    const colors = ['d9d2e9','f4cccc','d9ead3','fff2cc','d0dfe3']
    const dateIndex = 1;
    const amountIndex = 3;
    const startRowIndex = 2;
    const summaryIndex = 6;

    /* Parse Data */
    worksheet.spliceColumns(1,1);
    worksheet.spliceColumns(4,3);
    worksheet.spliceColumns(4,0,[],[],[]);
    worksheet.getRow(startRowIndex - 1).getCell(amountIndex + 1).value = 'remove'
    var colorIndex = 0;
    const firstRow = worksheet.getRow(startRowIndex);
    const firstRowDateStr =  firstRow.getCell(dateIndex).value;
    firstRow.getCell(amountIndex + 1).value = false;
    firstRow.getCell(dateIndex + 1).value = firstRow.getCell(dateIndex + 1).text.substring(0,19);
    var lastDate = dayjs(firstRowDateStr, "MM/DD/YYYY");
    firstRow.fill = {
        type: 'pattern',
        pattern: 'solid',
        fgColor: {argb: colors[colorIndex]}
    }
    var weekStartRow = startRowIndex;

    for(let i = startRowIndex + 1; i <= worksheet.rowCount; i++) {
        const row = worksheet.getRow(i);
        const date = row.getCell(dateIndex).value
        const dateDAYJS = dayjs(date, "MM/DD/YYYY");

        // check if not in the same week
        if(!dateDAYJS.isSame(lastDate, 'week')) {
            colorIndex = (colorIndex + 1) % colors.length;
            const label_cell = worksheet.getRow(weekStartRow).getCell(summaryIndex).value = "Week Summary";
            label_cell.fill = {type: 'pattern', pattern: 'solid', fgColor: {argb: "ffff00"}}
            label_cell.border = {
                top: {style:'thin'},
                left: {style:'thin'},
                bottom: {style:'thick'},
                right: {style:'thin'}
            };
            const s_cell = worksheet.getRow(weekStartRow + 1).getCell(summaryIndex);
            s_cell.value = { formula: 
                `SUMIFS(C${weekStartRow}:C${i-1},C${weekStartRow}:C${i-1},"<0",D${weekStartRow}:D${i-1},false)`
            }
            s_cell.fill = {type: 'pattern', pattern: 'solid', fgColor: {argb: "ffff00"}}
            s_cell.border = {
                top: {style:'thick'},
                left: {style:'thin'},
                bottom: {style:'thin'},
                right: {style:'thin'}
            };
            weekStartRow = i;
        }
        lastDate = dateDAYJS

        row.getCell(amountIndex + 1).value = false;
        row.getCell(dateIndex + 1).value = row.getCell(dateIndex + 1).text.substring(0,19);
        row.fill = {
            type: 'pattern',
            pattern: 'solid',
            fgColor: {argb: colors[colorIndex]}
        }
    }
    worksheet.columns.forEach(function (column, i) {
        let maxLength = 0;
        column["eachCell"]({ includeEmpty: true }, function (cell) {
            var columnLength = cell.value ? cell.value.toString().length : 10;
            if (columnLength > maxLength ) {
                maxLength = columnLength;
            }
        });
        column.width = maxLength < 10 ? 10 : maxLength;
    });
    console.log('##### Writing file to ' + process.argv[3]);
    await workbook.xlsx.writeFile(process.argv[3]);
}

const checkArgs = () => {
    if(process.argv.length <= 2) {
        console.log('Need to provide a file to read from. Must be a CSV file')
        process.exit();
    }
    if(process.argv.length <= 3) {
        console.log('Must provide an output file. File type of .xlsx')
        process.exit();
    }
    const argv1 = process.argv[2];
    const argv2 = process.argv[3];
    if(argv1.length - argv1.lastIndexOf(".CSV") != 4) {
        console.log('Input file parameter must be of type .CSV')
        process.exit();
    }
    if(argv2.length - argv2.lastIndexOf(".xlsx") != 5) {
        console.log('Output file parameter must be of type .xlsx')
        process.exit();
    }
    console.log(`##### Parsing File: ${argv1}`)
}

checkArgs();
main();
