const { readFileSync, writeFileSync } = require('fs')
const { parseXlsx } = require('..')

; (async () => {
  const xlsxFile = readFileSync('./test/test.xlsx')
  const worksheets = await parseXlsx(xlsxFile) // or parseXlsxFile('./test/dimibob.xlsx')

  console.log('%d sheet(s) found', worksheets.size)

  console.log('\nI\'ll use `sheet1` sheet')
  const sheet = worksheets.get('sheet1')

  console.log('ㄴ there are  %d rows', sheet.rows.size)
  console.log('ㄴ there are %d cols', sheet.rows.size * sheet.rows.get('1').size)

  console.log('\nContent of `A1` cell:', sheet.getCell('A1').value)

  console.log('\nI\'ll create `test/array.json` file for testing toArray()')
  writeFileSync('./test/array.json', JSON.stringify(sheet.toArray(), null, 2))

  console.log('I\'ll create `test/flatedArray.json` file for testing toFlatArray()')
  writeFileSync('./test/flatedArray.json', JSON.stringify(sheet.toFlatArray(), null, 2))
})()
