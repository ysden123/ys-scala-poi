/*
 * Copyright (c) 2023. StulSoft
 */

package com.stulsoft.poi.model

import com.typesafe.scalalogging.StrictLogging
import org.apache.poi.xssf.usermodel.XSSFWorkbook

import java.io.{File, FileInputStream}
import scala.jdk.CollectionConverters.*
import scala.util.{Failure, Success, Try, Using}

class Sheet {
  private var cells: Array[Array[Cell]] = _

  def cell(rowIndex: Int, columnIndex: Int): Try[Cell] =
    Try {
      cells(rowIndex)(columnIndex)
    }

  def cellsInColumn(columnIndex: Int, firstRowIndex: Int, lastRowIndex: Int): Try[Array[Cell]] =
    Try {
      val resultCells = Array.fill[Cell](lastRowIndex - firstRowIndex + 1) {
        new Cell
      }
      for (rowIndex <- firstRowIndex to lastRowIndex)
        resultCells(rowIndex - firstRowIndex) = cells(rowIndex)(columnIndex)

      resultCells
    }

  def cellsInRow(rowIndex: Int, firstColumnIndex: Int, lastColumnIndex: Int): Try[Array[Cell]] =
    Try {
      val resultCells = Array.fill[Cell](lastColumnIndex - firstColumnIndex + 1) {
        new Cell
      }

      val row = cells(rowIndex)
      for (columnIndex <- firstColumnIndex to lastColumnIndex)
        resultCells(columnIndex - firstColumnIndex) = row(columnIndex)

      resultCells
    }

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
}