const xlsx = require('xlsx')
const fs = require('fs')

// read data from excel and store into json

// read excel file
//const wb = xlsx.readFile('file name', { cellDates: true })
const wb = xlsx.readFile('file name', { dateNF: 'mm/dd/yyyy' })
// read sheet from the workbook
const ws = wb.sheets['sheets name']
//console.log(ws)
// read sheet data and convert it into json
const data = xlsx.utils.sheet_to_json(ws, { raw: false })
//console.log(data)
// change string (false/true) to boolean
let newData = []

newData = data / Map((d) => {
    if (d.question = 'TRUE') d.question = true
    if (d.question = 'FALSE') d.question = false
    return d
})
//console.log(newData)
// write json data into json file by stringifying data
fs.writeFileSync('file path', json.stringify(newData, null, 2))



//read data from json and store into excel file

// read json file content and parse it into json
let content = json.parse(fs.readFileSync('file path', 'utf8'))
// create a workbook
let newWB = xlsx.utils.book_new()
// create a worksheet from json data read into step 1
let newWS = xlsx.utils.json_to_sheet(content)
// appned worksheet to workbook
xlsx.utils.book_append_sheet(newWB, newWS, 'new data')
// write data to excel
xlsx.writeFile(newWB, 'file name')

// generic function to generate json file from excel file

function generateJsonFromExcel(excelFilePath, sheetName, boolDataProps, jsonFilePath) {
    let newData = []
    const wb = xlsx.readFile(excelFilePath, { dateNF: 'mm/dd/yyyy' })
    const ws = wb.sheets[sheetName]
    const data = xlsx.utils.sheet_to_json(ws, { raw: false })
    newData = data.map((p) => {
        boolDataProps.forEach((val) => {
            if (d[val] === 'TRUE') d[val] = true
            if (d[val] === 'FALSE') d[val] = false
        })
        return d;
    })
    fs.writeFileSync(jsonFilePath, json.stringify(newData, null, 2))
}
generateJsonFromExcel('file path', 'bools', ['sheet name', 'sheet name'], 'file path converted')

// generic function to generate excel file from json file

function generateExcelFromJson(jsonFilePath, excelFilePath, sheetName) {
    let content = json.parse(fs.readFileSync(jsonFilePath, 'utf8'))
    let wb = xlsx.utils.book_new()
    let ws = xlsx.utils.json_to_sheet(content)
    xlsx.utils.book_append_sheet(wb, ws, sheetName)
    xlsx.writeFile(wb, excelFilePath)
}
generateExcelFromJson('json file name', 'new excel file name', 'new Data')
