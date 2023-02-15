/*
 * Copyright (c) 2023. StulSoft
 */

package com.stulsoft.poi.model

enum CellType {
  case BLANK, BOOLEAN, ERROR, FORMULA, NUMERIC, STRING
}

object CellType{
  def valueOf(javaCellType: org.apache.poi.ss.usermodel.CellType): CellType=
    CellType.valueOf(javaCellType.name())
}