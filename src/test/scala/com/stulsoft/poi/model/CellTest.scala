/*
 * Copyright (c) 2023. StulSoft
 */

package com.stulsoft.poi.model

import org.apache.poi.xssf.streaming.{SXSSFCell, SXSSFRow, SXSSFSheet, SXSSFWorkbook}
import org.apache.poi.xssf.usermodel.{XSSFCell, XSSFSheet}
import org.scalatest.funsuite.AnyFunSuite

import java.time.LocalDateTime

class CellTest extends AnyFunSuite {
  test("cellType") {
    val cell = new Cell
    assert(CellType.BLANK == cell.cellType)

    cell.cellType = CellType.NUMERIC
    assert(CellType.NUMERIC == cell.cellType)
  }

  test("constructor") {
    try {
      val workbook: SXSSFWorkbook = new SXSSFWorkbook()
      val sheet: SXSSFSheet = new SXSSFSheet(workbook, null)
      val row: SXSSFRow = new SXSSFRow(sheet)
      row.setRowNum(0)

      val javaCell = new SXSSFCell(row, org.apache.poi.ss.usermodel.CellType.NUMERIC)
      javaCell.setCellValue(1.234)

      var cell: Cell = Cell(javaCell)
      assert(CellType.NUMERIC == cell.cellType)
      assert(1.234 == cell.cellDoubleValue())

      javaCell.setCellValue("test")
      cell = Cell(javaCell)
      assert(CellType.STRING == cell.cellType)
      assert("test" == cell.cellStringValue())
    } catch
      case exception: Exception => exception.printStackTrace()
  }
}
