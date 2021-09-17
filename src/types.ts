import { Worksheet } from './classes/Worksheet'

/**
 * 하나의 셸의 대한 정보입니다.
 */
export interface CellData {
  /** 해당 셸의 값입니다. */
  value: string
  /** 해당 셸의 자료형입니다. `s`인 경우 `mappingShardedStrings()`를 진행해야 합니다. */
  type: 's' | 'n'
}

/**
 * 해당 행의 번호입니다. (ex: 1, 2, 3, 4...)
 */
export type RowNn = string

/**
 * 해당 셸의 번호입니다. (ex: A1, B2, C3...)
 */
export type CellNn = string

/**
 * 시트 이름입니다. (ex: sheet1)
 */
export type SheetName = string

/**
 * 시트 목록입니다
 */
export type Worksheets = Map<SheetName, Worksheet>

/**
 * ShardedStrings 내부파일의 대한 정보입니다.
 */
export type ShardedStringsData = string[]

/**
 * xlsx 파일의 대한 가공되지 않은 원시 정보입니다.
 */
export interface XLSXRawData {
  sheets: Map<SheetName, string>,
  sharedStrings: string
}
