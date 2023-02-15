/*
 * Copyright (c) 2023. StulSoft
 */

package com.stulsoft.poi.model


import scala.reflect.ClassTag.Nothing

trait CellValue

class CellDoubleValue(val value:Double) extends CellValue
class CellStringValue(val value:String) extends CellValue
class CellBooleanValue(val value:Boolean) extends CellValue

class Cell {
  var cellType: CellType = CellType.BLANK
  private var cellValue: CellValue = _
  def cellDoubleValue():Double = cellValue.asInstanceOf[CellDoubleValue].value
  def cellStringValue():String = cellValue.asInstanceOf[CellStringValue].value
  def cellBooleanValue():Boolean = cellValue.asInstanceOf[CellBooleanValue].value
}

object Cell {
  def apply(cell: org.apache.poi.ss.usermodel.Cell): Cell =
    val newCell = new Cell
    newCell.cellType = CellType.valueOf(cell.getCellType)
    newCell.cellValue = newCell.cellType match
      case CellType.BLANK | CellType.ERROR => null
      case CellType.NUMERIC =>  new CellDoubleValue(cell.getNumericCellValue)
      case CellType.STRING => new CellStringValue(cell.getStringCellValue)
      case CellType.BOOLEAN => new CellBooleanValue(cell.getBooleanCellValue)
      case _ => null
    newCell
}
