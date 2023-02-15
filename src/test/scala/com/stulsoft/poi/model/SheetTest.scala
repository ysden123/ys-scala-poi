/*
 * Copyright (c) 2023. StulSoft
 */

package com.stulsoft.poi.model

import org.scalatest.funsuite.AnyFunSuite

import scala.util.{Failure, Success}

class SheetTest extends AnyFunSuite {
  test("constructor with wrong file location"){
    Sheet.build("ERROR.xlsx") match
      case Success(_) => fail("Exception is expected")
      case Failure(exception) =>
        println(exception.getMessage)
        succeed
  }

  test("constructor") {
    Sheet.build("src/test/resources/Test1.xlsx") match
      case Success(sheet) =>
        val cell1 = sheet.cell(0, 0)
        assert(cell1.isSuccess)
        assert(1.0 == cell1.get.cellDoubleValue())

        val cell2 = sheet.cell(0, 1)
        assert(cell2.isSuccess)
        assert("One" == cell2.get.cellStringValue())

        val cell3 = sheet.cell(2, 0)
        assert(cell3.isSuccess)
        assert(CellType.BLANK == cell3.get.cellType)

        val cellError = sheet.cell(10, 10)
        assert(cellError.isFailure)
        cellError match
          case Failure(_) =>
            succeed
          case Success(_) =>
            fail("Should be exception")
      case Failure(exception) => fail(exception.getMessage)
  }

  test("cellsInColumn") {
    Sheet.build("src/test/resources/Test1.xlsx") match
      case Success(sheet) =>
        val cells = sheet.cellsInColumn(0, 0, 1)
        assert(cells.isSuccess)
        val cellValues = cells.get
        assert(2 == cellValues.length)
        assert(1 == cellValues(0).cellDoubleValue())
      case Failure(exception) => fail(exception.getMessage)
  }

  test("cellsInColumn 2") {
    Sheet.build("src/test/resources/Test1.xlsx") match
      case Success(sheet) =>
        val cells = sheet.cellsInColumn(0, 1, 1)
        assert(cells.isSuccess)
        val cellValues = cells.get
        assert(1 == cellValues.length)
        assert(2 == cellValues(0).cellDoubleValue())
      case Failure(exception) => fail(exception.getMessage)
  }

  test("cellsInRow") {
    Sheet.build("src/test/resources/Test1.xlsx") match
      case Success(sheet) =>
        val cells = sheet.cellsInRow(1, 0, 1)
        assert(cells.isSuccess)
        val cellValues = cells.get
        assert(2 == cellValues.length)
        assert(2 == cellValues(0).cellDoubleValue())
        assert("Two" == cellValues(1).cellStringValue())
      case Failure(exception) => fail(exception.getMessage)
  }

  test("sumInColumn") {
    Sheet.build("src/test/resources/Test1.xlsx") match
      case Success(sheet) =>
        sheet.sumInColumn(0, 0, 3) match
          case Success(sum) =>
            assert(sum == 3.0)
          case Failure(exception) =>
            fail(exception.getMessage)
      case Failure(exception) => fail(exception.getMessage)
  }

  test("sumInRow") {
    Sheet.build("src/test/resources/Test1.xlsx") match
      case Success(sheet) =>
        sheet.sumInRow(0, 0, 1) match
          case Success(sum) =>
            assert(sum == 1.0)
          case Failure(exception) =>
            fail(exception.getMessage)

        sheet.sumInRow(1, 0, 1) match
          case Success(sum) =>
            assert(sum == 2.0)
          case Failure(exception) =>
            fail(exception.getMessage)

        sheet.sumInRow(1, 1, 1) match
          case Success(sum) =>
            assert(sum == 0)
          case Failure(exception) =>
            fail(exception.getMessage)
      case Failure(exception) => fail(exception.getMessage)
  }

  test("foldLeftInColumn") {
    Sheet.build("src/test/resources/Test1.xlsx") match
      case Success(sheet) =>
        sheet
//        .foldLeftInColumn(0, 0, 3, CellType.NUMERIC, 0.0)(_ + _.cellDoubleValue()) match
//        .foldLeftInColumn(0, 0, 3, CellType.NUMERIC, 0.0)((acc: Double, cell: Cell) => acc + cell.cellDoubleValue()) match
          .foldLeftInColumn(0, 0, 3, CellType.NUMERIC, 0.0)((acc, cell) => acc + cell.cellDoubleValue()) match
          case Success(sum) =>
            assert(sum == 3.0)
          case Failure(exception) =>
            fail(exception.getMessage)

        sheet
          .foldLeftInColumn(1, 0, 3, CellType.STRING, "")((acc, cell) => acc + cell.cellStringValue()) match
          case Success(concatenated) =>
            assert(concatenated == "OneTwo")
          case Failure(exception) =>
            fail(exception.getMessage)
      case Failure(exception) => fail(exception.getMessage)
  }
}
