// Require library
var xl = require('excel4node');
const testData = [{ name: "me", color: "red", some: "lol", red: "lol", small: "eeee" }, { name: "chris", color: "red", some: "lol", red: "lol", small: "eeee" }, { name: "chris", color: "red", some: "lol", red: "lol", small: "eeee" }, { name: "chris", color: "red", some: "lol", red: "lol", small: "eeee" }, { name: "chris", color: "red", some: "lol", red: "lol", small: "eeee" }, { name: "chris", color: "red", some: "lol", red: "lol", small: "eeee" }, { name: "chris", color: "red", some: "lol", red: "lol", small: "eeee" }, { name: "chris", color: "red", some: "lol", red: "lol", small: "eeee" }, { name: "chris", color: "red", some: "lol", red: "lol", small: "eeee" }, { name: "chris", color: "red", some: "lol", red: "lol", small: "eeee" }, { name: "chris", color: "red", some: "lol", red: "lol", small: "eeee" }, { name: "chris", color: "red", some: "lol", red: "lol", small: "eeee" }, { name: "chris", color: "red", some: "lol", red: "lol", small: "eeee" }, { name: "chris", color: "red", some: "lol", red: "lol", small: "eeee" }, { name: "chris", color: "red", some: "lol", red: "lol", small: "eeee" }, { name: "chris", color: "red", some: "lol", red: "lol", small: "eeee" }, { name: "chris", color: "red", some: "lol", red: "lol", small: "eeee" }, { name: "chris", color: "red", some: "lol", red: "lol", small: "eeee" }, { name: "chris", color: "red", some: "lol", red: "lol", small: "eeee" }, { name: "chris", color: "red", some: "lol", red: "lol", small: "eeee" }, { name: "chris", color: "red", some: "lol", red: "lol", small: "eeee" }, { name: "chris", color: "red", some: "lol", red: "lol", small: "eeee" }, { name: "chris", color: "red", some: "lol", red: "lol", small: "eeee" }, { name: "chris", color: "red", some: "lol", red: "lol", small: "eeee" }, { name: "chris", color: "red", some: "lol", red: "lol", small: "eeee" }, { name: "chris", color: "red", some: "lol", red: "lol", small: "eeee" }, { name: "chris", color: "red", some: "lol", red: "lol", small: "eeee" }, { name: "chris", color: "red", some: "lol", red: "lol", small: "eeee" }, { name: "chris", color: "red", some: "lol", red: "lol", small: "eeee" },]

// Create a new instance of a Workbook class
function generateExcel(data) {
    var wb = new xl.Workbook();

    // Add Worksheets to the workbook
    var ws = wb.addWorksheet('Sheet 1');
    var ws2 = wb.addWorksheet('Sheet 2');

    // Create a reusable style
    var style = wb.createStyle({
        font: {
            color: '#000103',
            size: 12,
        }
    });

    const title = Object.keys(testData[0])
    let values = []
    testData.forEach((el) => {
        values.push(...Object.values(el))
    })


    let titleCount = 1
    let colCount = 2
    let rowCount = 1
    let arrPosition = title.length

    for (let j = 0; j < title.length; j++) {
        ws.cell(1, titleCount).string(title[j])

        for (let i = 0; i < values.length; i++) {
            if (i % arrPosition === 0) {
                ws.cell(colCount, rowCount).string(values[i]).style(style)
                colCount++
            };

        }
        colCount = 2
        titleCount++
        rowCount++
        values.shift()
    }





    wb.write('Excel.xlsx');
}

generateExcel()