import JSZip from 'jszip'
import { ElementCompact } from 'xml-js'
import { Worksheet } from './classes/Worksheet'
import { ShardedStringsData, SheetName, XLSXRawData } from './types'

export function ifElemName (name: string) {
  return (elem: ElementCompact) => elem.name === name
}

export async function extract (stream: any): Promise<XLSXRawData> {
  const data = await JSZip.loadAsync(stream)
  const worksheets = data.filter((path) => path.startsWith('xl/worksheets/'))
  const sharedStrings = await data.file('xl/sharedStrings.xml')?.async('string')

  if (!sharedStrings) throw new Error('Shared strings not found')

  const sheets = new Map<SheetName, string>()
  for (const rawSheet of Object.values(worksheets)) {
    sheets.set(rawSheet.name, await rawSheet.async('string'))
  }

  return { sheets, sharedStrings }
}

export function mappingShardedStrings (sheet: Worksheet, sharedStrings: ShardedStringsData) {
  for (const [rowKey, rowValue] of sheet.rows.entries()) {
    for (const [cellKey, cellValue] of rowValue.entries()) {
      if (cellValue.type !== 's') continue
      cellValue.value = sharedStrings[parseInt(cellValue.value)]

      rowValue.set(cellKey, cellValue)
      sheet.rows.set(rowKey, rowValue)
    }
  }

  return sheet
}
