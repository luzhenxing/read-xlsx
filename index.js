const xlsx = require('xlsx')

const workbook = xlsx.readFile('./doc/rank.xlsx', { type: 'utf-8' })

const sheetNames = workbook.SheetNames

const sheetTab = workbook.Sheets[sheetNames[0]]

const sheetJson = xlsx.utils.sheet_to_json(sheetTab)

sheetJson.forEach(item => {
    console.log(item)
})