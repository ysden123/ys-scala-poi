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
        val cell1 = sheet.cell(1, 1)
        assert(cell1.isSuccess)
        assert(1.0 == cell1.get.cellDoubleValue())

        val cell2 = sheet.cell(2,1)
        assert(cell2.isSuccess)
        assert("One" == cell2.get.cellStringValue())

        val cell3 = sheet.cell(1,3)
        assert(cell3.isSuccess)
        assert(CellType.BLANK == cell3.get.cellType)

        val cellError = sheet.cell(11, 11)
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
        val cells = sheet.cellsInColumn(1, 1, 2)
        assert(cells.isSuccess)
        val cellValues = cells.get
        assert(2 == cellValues.length)
        assert(1 == cellValues(0).cellDoubleValue())
      case Failure(exception) => fail(exception.getMessage)
  }

  test("cellsInColumn 2") {
    Sheet.build("src/test/resources/Test1.xlsx") match
      case Success(sheet) =>
        val cells = sheet.cellsInColumn(1, 2, 2)
        assert(cells.isSuccess)
        val cellValues = cells.get
        assert(1 == cellValues.length)
        assert(2 == cellValues(0).cellDoubleValue())
      case Failure(exception) => fail(exception.getMessage)
  }

  test("cellsInRow") {
    Sheet.build("src/test/resources/Test1.xlsx") match
      case Success(sheet) =>
        val cells = sheet.cellsInRow(2, 1, 2)
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
        sheet.sumInColumn(1, 1, 4) match
          case Success(sum) =>
            assert(sum == 3.0)
          case Failure(exception) =>
            fail(exception.getMessage)
      case Failure(exception) => fail(exception.getMessage)
  }

  test("sumInRow") {
    Sheet.build("src/test/resources/Test1.xlsx") match
      case Success(sheet) =>
        sheet.sumInRow(1, 1, 2) match
          case Success(sum) =>
            assert(sum == 1.0)
          case Failure(exception) =>
            fail(exception.getMessage)

        sheet.sumInRow(2, 1, 2) match
          case Success(sum) =>
            assert(sum == 2.0)
          case Failure(exception) =>
            fail(exception.getMessage)

        sheet.sumInRow(2, 2, 2) match
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
//        .foldLeftInColumn(1, 1, 4, CellType.NUMERIC, 0.0)(_ + _.cellDoubleValue()) match
//        .foldLeftInColumn(1, 1, 4, CellType.NUMERIC, 0.0)((acc: Double, cell: Cell) => acc + cell.cellDoubleValue()) match
          .foldLeftInColumn(1, 1, 4, CellType.NUMERIC, 0.0)((acc, cell) => acc + cell.cellDoubleValue()) match
          case Success(sum) =>
            assert(sum == 3.0)
          case Failure(exception) =>
            fail(exception.getMessage)

        sheet
          .foldLeftInColumn(2, 1, 4, CellType.STRING, "")((acc, cell) => acc + cell.cellStringValue()) match
          case Success(concatenated) =>
            assert(concatenated == "OneTwo")
          case Failure(exception) =>
            fail(exception.getMessage)
      case Failure(exception) => fail(exception.getMessage)
  }

  test("columnNumber"){
    var index = Sheet.columnNumber("")
    assert(index == 0)
    index = Sheet.columnNumber("a")
    assert(index == 1)
    index = Sheet.columnNumber("B")
    assert(index == 2)

    assert(Sheet.columnNumber("A") == 1)
    assert(Sheet.columnNumber("Z") == 26)
    assert(Sheet.columnNumber("AA") == 27)
    assert(Sheet.columnNumber("AB") == 28)
    assert(Sheet.columnNumber("AAB") == 704)
    assert(Sheet.columnNumber("ABB") == 730)
    assert(Sheet.columnNumber("AZ") == 52)
    assert(Sheet.columnNumber("BA") == 53)
    assert(Sheet.columnNumber("BB") == 54)
    assert(Sheet.columnNumber("bBb") == 1406)
    assert(Sheet.columnNumber("AAa") == 703)
  }

  test("cell"){
    Sheet.build("src/test/resources/Test1.xlsx") match
      case Success(sheet) =>
        sheet.cell("a",1) match
          case Success(cell) =>
            assert(cell.cellDoubleValue() == 1.0)
          case Failure(exception) =>
            fail(exception.getMessage)
      case Failure(exception) => fail(exception.getMessage)
  }
}
