import { Blob } from 'buffer'
import { readFileSync } from 'fs'

import { Worksheet } from './classes/Worksheet'
import { SheetName, Worksheets } from './types'
import { extract, mappingShardedStrings } from './utils'
import { parseSharedStringsXml, parseSheetXml } from './parsers'

/**
 * xlsx 데이터를 파싱합니다.
 *
 * @param input xlsx파일의 2진수 데이터 (Buffer, Blob, number배열 모두 가능)
 */
export async function parseXlsx (input: Buffer | ArrayBuffer | Blob | Uint8Array | number[]) {
  const extracted = await extract(input)
  const sharedString = parseSharedStringsXml(extracted.sharedStrings)

  const worksheets = new Map<SheetName, Worksheet>()
  for (const [rawSheetKey, rawSheetValue] of extracted.sheets) {
    if (!rawSheetValue.includes('<worksheet') && !rawSheetValue.includes('<x:worksheet')) continue

    const worksheet = new Worksheet(parseSheetXml(rawSheetValue))
    worksheets.set(rawSheetKey.replace(/xl\/worksheets\/|.xml$/g, ''),
      mappingShardedStrings(worksheet, sharedString))
  }

  return worksheets as Worksheets
}

/**
 * xlsx 파일을 파싱합니다.
 *
 * @param path xlsx파일 경로
 */
export function parseXlsxFile (path: string) {
  return parseXlsx(readFileSync(path))
}
