import { CellData, CellNn, RowNn } from '../types'

export class Worksheet {
  /**
   * 시트의 행 목록입니다.
   * Key: 행번호 (ex: 1, 2, 3, 4...)
   * Value: 행의 셸 정보들
   */
  public rows: Map<RowNn, Map<CellNn, CellData>>

  constructor (rows: Map<RowNn, Map<CellNn, CellData>>) {
    this.rows = rows
  }

  /**
   * 시트에서 해당 번호의 셸을 반환합니다.
   *
   * @example
   *   // A1번의 셸 값 불러오기
   *   const cell = worksheet.getCell('A1')
   *   console.log(cell.value)
   *
   * @param key 셸 번호 (ex: A1, B2)
   */
  public getCell (key: string) {
    if (key.length !== 2) throw new Error(`invalid cell number '${key}', (length too short)`)
    return this.rows.get(key.charAt(1))?.get(key)
  }

  /**
   * 시트에서 해당 번호의 행을 반환합니다.
   *
   * @example
   *  // 1번째 행의 셸 개수 불러오기
   *  const row = worksheet.getRow('1')
   *  console.log(row.length)
   *
   * @param key 행 번호
   */
  public getRow (key: string) {
    return this.rows.get(key)
  }

  /**
   * 해당 시트를 Flat 시킵니다.
   * 즉, 모든 값을 1차원 배열로 반환합니다.
   */
  public toFlatArray () {
    const data = [] as string[]

    for (const [, row] of this.rows) {
      for (const [, cell] of row) {
        data.push(cell.value)
      }
    }

    return data
  }

  /**
   * 해당 시트를 배열로 변환 시킵니다.
   * 행, 열을 살린 2차원 배열로 반환합니다.
   */
  public toArray () {
    const datas = [] as string[][]

    for (const [, row] of this.rows) {
      const data = [] as string[]
      for (const [, cell] of row) {
        data.push(cell.value)
      }
      datas.push(data)
    }

    return datas
  }
}
