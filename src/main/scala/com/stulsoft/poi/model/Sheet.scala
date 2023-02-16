/*
 * Copyright (c) 2023. StulSoft
 */

package com.stulsoft.poi.model

import com.stulsoft.poi.model.Sheet.columnNumber
import com.typesafe.scalalogging.StrictLogging
import org.apache.poi.xssf.usermodel.XSSFWorkbook

import java.io.{File, FileInputStream}
import scala.jdk.CollectionConverters.*
import scala.util.{Failure, Success, Try, Using}

class Sheet {
  private var cells: Array[Array[Cell]] = _

  /**
   * Returns a cell for specified column and row.
   *
   * @param columnIndex the column index, starts from 1
   * @param rowIndex    the row index, starts from 1
   * @return the cell for specified column and row.
   */
  def cell(columnIndex: Int, rowIndex: Int): Try[Cell] =
    Try {
      cells(rowIndex - 1)(columnIndex - 1)
    }

  /**
   * Returns a cell for specified column and row.
   *
   * @param columnIndex the column index, starts from 'A'
   * @param rowIndex    the row index, starts from 1
   * @return the cell for specified column and row.
   */
  def cell(columnIndex: String, rowIndex: Int): Try[Cell] =
    cell(columnNumber(columnIndex), rowIndex)

  /**
   * Returns an array of the cells in the specified column
   *
   * @param columnIndex   specifies column, starts from 1
   * @param firstRowIndex number of the first row, starts from 1
   * @param lastRowIndex  number of the last row, starts from 1
   * @return the array of the cells in the specified column
   */
  def cellsInColumn(columnIndex: Int, firstRowIndex: Int, lastRowIndex: Int): Try[Array[Cell]] =
    Try {
      val resultCells = Array.fill[Cell](lastRowIndex - firstRowIndex + 1) {
        new Cell
      }
      for (rowIndex <- firstRowIndex to lastRowIndex)
        resultCells(rowIndex - firstRowIndex) = cells(rowIndex - 1)(columnIndex - 1)

      resultCells
    }

  /**
   * Returns an array of the cells in the specified column.
   *
   * @param columnIndex   specifies column, starts from 'A'
   * @param firstRowIndex number of the first row, starts from 1
   * @param lastRowIndex  number of the last row, starts from 1
   * @return the array of the cells in the specified column
   */
  def cellsInColumn(columnIndex: String, firstRowIndex: Int, lastRowIndex: Int): Try[Array[Cell]] =
    cellsInColumn(columnNumber(columnIndex), firstRowIndex, lastRowIndex)

  /**
   * Returns an array of the cells in the specified row.
   *
   * @param rowIndex         specifies the row, starts from 1
   * @param firstColumnIndex number of the first column, starts from 1
   * @param lastColumnIndex  number of the last column, starts from 1
   * @return the an array of the cells in the specified row
   */
  def cellsInRow(rowIndex: Int, firstColumnIndex: Int, lastColumnIndex: Int): Try[Array[Cell]] =
    Try {
      val resultCells = Array.fill[Cell](lastColumnIndex - firstColumnIndex + 1) {
        new Cell
      }

      val row = cells(rowIndex - 1)
      for (columnIndex <- firstColumnIndex to lastColumnIndex)
        resultCells(columnIndex - firstColumnIndex) = row(columnIndex - 1)

      resultCells
    }

  /**
   * Returns an array of the cells in the specified row.
   *
   * @param rowIndex         specifies the row, starts from 1
   * @param firstColumnIndex number of the first column, starts from 'A'
   * @param lastColumnIndex  number of the last column, starts from 'A'
   * @return the an array of the cells in the specified row
   */
  def cellsInRow(rowIndex: Int, firstColumnIndex: String, lastColumnIndex: String): Try[Array[Cell]] =
    cellsInRow(rowIndex, columnNumber(firstColumnIndex), columnNumber(lastColumnIndex))

  /**
   * Calculates a sum of the cells in the specified column.
   *
   * @param columnIndex   specifies the column, starts from 1
   * @param firstRowIndex specifies the first row, starts from 1
   * @param lastRowIndex  specifies the last row, starts from 1
   * @return the sum of the cells in the specified column
   */
  def sumInColumn(columnIndex: Int, firstRowIndex: Int, lastRowIndex: Int): Try[Double] =
    Try {
      cellsInColumn(columnIndex, firstRowIndex, lastRowIndex) match
        case Success(selectedCells) =>
          selectedCells
            .toList
            .filter(cell => cell.cellType == CellType.NUMERIC)
            .map(cell => cell.cellDoubleValue())
            .sum
        case Failure(exception) =>
          throw exception
    }

  /**
   * Calculates a sum of the cells in the specified column.
   *
   * @param columnIndex   specifies the column, starts from 'A'
   * @param firstRowIndex specifies the first row, starts from 1
   * @param lastRowIndex  specifies the last row, starts from 1
   * @return the sum of the cells in the specified column
   */
  def sumInColumn(columnIndex: String, firstRowIndex: Int, lastRowIndex: Int): Try[Double] =
    sumInColumn(columnNumber(columnIndex), firstRowIndex, lastRowIndex)

  /**
   * Calculates a sum of the cells in the specified row.
   *
   * @param rowIndex         specifies the  row, starts from 1
   * @param firstColumnIndex specifies the first column, starts from 1
   * @param lastColumnIndex  specifies the last column, starts from 1
   * @return the sum of the cells in the specified row.
   */
  def sumInRow(rowIndex: Int, firstColumnIndex: Int, lastColumnIndex: Int): Try[Double] =
    Try {
      cellsInRow(rowIndex, firstColumnIndex, lastColumnIndex) match
        case Success(selectedCells) =>
          selectedCells
            .toList
            .filter(cell => cell.cellType == CellType.NUMERIC)
            .map(cell => cell.cellDoubleValue())
            .sum
        case Failure(exception) =>
          throw exception
    }

  /**
   * Calculates a sum of the cells in the specified row.
   *
   * @param rowIndex         specifies the  row, starts from 1
   * @param firstColumnIndex specifies the first column, starts from 'A'
   * @param lastColumnIndex  specifies the last column, starts from 'A'
   * @return the sum of the cells in the specified row.
   */
  def sumInRow(rowIndex: Int, firstColumnIndex: String, lastColumnIndex: String): Try[Double] =
    sumInRow(rowIndex, columnNumber(firstColumnIndex), columnNumber(lastColumnIndex))

  /**
   * Applies the specified operation for the cells in the specified column
   *
   * @param columnIndex   specifies the column, starts from 1
   * @param firstRowIndex specifies the first row, starts from 1
   * @param lastRowIndex  specifies the last row, starts from 1
   * @param cellType      specifies the cell type; the operation will be applied for these cells only
   * @param z             the initial value
   * @param op            the operation
   * @tparam B type of the result
   * @return the result of operation applied on cells
   */
  def foldLeftInColumn[B](columnIndex: Int, firstRowIndex: Int, lastRowIndex: Int, cellType: CellType, z: B)(op: (B, Cell) => B): Try[B] =
    Try {
      cellsInColumn(columnIndex, firstRowIndex, lastRowIndex) match
        case Success(selectedCells) =>
          selectedCells
            .toList
            .filter(cell => cell.cellType == cellType)
            .foldLeft(z)(op)
        case Failure(exception) =>
          throw exception
    }

  /**
   * Applies the specified operation for the cells in the specified column
   *
   * @param columnIndex   specifies the column, starts from 'A'
   * @param firstRowIndex specifies the first row, starts from 1
   * @param lastRowIndex  specifies the last row, starts from 1
   * @param cellType      specifies the cell type; the operation will be applied for these cells only
   * @param z             the initial value
   * @param op            the operation
   * @tparam B type of the result
   * @return the result of operation applied on cells
   */
  def foldLeftInColumn[B](columnIndex: String, firstRowIndex: Int, lastRowIndex: Int, cellType: CellType, z: B)(op: (B, Cell) => B): Try[B] =
    foldLeftInColumn(columnNumber(columnIndex), firstRowIndex, lastRowIndex, cellType, z)(op)
}

object Sheet extends StrictLogging {
  def build(fileLocation: String): Try[Sheet] =
    Try {
      Using(new FileInputStream(new File(fileLocation))) {
        stream => {
          val sheet = new Sheet
          val workbook = new XSSFWorkbook(stream)
          val originalSheet = workbook.getSheetAt(0).asScala
          // Define dimension
          var rowNumber = -1
          var columnNumber = 0
          for (row <- originalSheet) {
            rowNumber = Math.max(rowNumber, row.getRowNum)
            columnNumber = Math.max(columnNumber, row.asScala.size)
          }

          sheet.cells = Array.fill[Cell](rowNumber + 1, columnNumber) {
            new Cell
          }

          for (row <- originalSheet) {
            val rowIndex = row.getRowNum
            for (originalCell <- row.asScala) {
              val columnIndex = originalCell.getColumnIndex
              sheet.cells(rowIndex)(columnIndex) = Cell(originalCell)
            }
          }

          sheet
        }
      } match
        case Success(sheet) => sheet
        case Failure(exception) => throw exception
    }

  def columnNumber(text: String): Int =
    val base = 'Z' - 'A' + 1
    var index = 0
    for (i <- 0 until text.length) {
      val power = text.length - i - 1
      index += Math.pow(base, power).toInt * (text.charAt(i).toUpper - 'A' + 1)
    }

    index
}