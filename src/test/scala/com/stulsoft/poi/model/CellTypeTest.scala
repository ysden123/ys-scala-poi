/*
 * Copyright (c) 2023. StulSoft
 */

package com.stulsoft.poi.model

import com.stulsoft.poi.model.CellType
import org.scalatest.funsuite.AnyFunSuite

class CellTypeTest extends AnyFunSuite {
  test("from java cell types") {
    assert(CellType.valueOf(org.apache.poi.ss.usermodel.CellType.BLANK) == CellType.BLANK)
    assert(CellType.valueOf(org.apache.poi.ss.usermodel.CellType.BOOLEAN) == CellType.BOOLEAN)
    assert(CellType.valueOf(org.apache.poi.ss.usermodel.CellType.ERROR) == CellType.ERROR)
    assert(CellType.valueOf(org.apache.poi.ss.usermodel.CellType.FORMULA) == CellType.FORMULA)
    assert(CellType.valueOf(org.apache.poi.ss.usermodel.CellType.NUMERIC) == CellType.NUMERIC)
    assert(CellType.valueOf(org.apache.poi.ss.usermodel.CellType.STRING) == CellType.STRING)
  }
}
