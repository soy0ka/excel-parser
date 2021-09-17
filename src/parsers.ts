import { ifElemName } from './utils'
import { ElementCompact, xml2js } from 'xml-js'
import { CellData, CellNn, RowNn, ShardedStringsData } from './types'

export function parseSheetXml (rawString: string) {
  const xml = xml2js(rawString)
  const rows = xml
    .elements.find(ifElemName('worksheet'))
    .elements.find(ifElemName('sheetData'))
    .elements.filter(ifElemName('row')) || []

  const data = new Map<RowNn, Map<CellNn, CellData>>()

  for (const row of rows as ElementCompact[]) {
    const rawCells = row.elements.filter(ifElemName('c')) as ElementCompact[]
    const cellData = new Map<CellNn, CellData>()

    for (const rawCell of rawCells) {
      cellData.set(rawCell.attributes.r, {
        type: rawCell.attributes.t,
        value: rawCell.elements?.find(ifElemName('v'))?.elements[0].text
      })
    }

    data.set(row.attributes.r, cellData)
  }

  return data
}

export function parseSharedStringsXml (rawString: string): ShardedStringsData {
  const xml = xml2js(rawString)
  const strings = xml
    .elements.find(ifElemName('sst'))
    .elements.filter(ifElemName('si'))
    .map((si: ElementCompact) =>
      si.elements.find(ifElemName('t'))?.elements[0].text)

  return strings
}
