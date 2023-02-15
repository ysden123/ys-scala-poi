/*
 * Copyright (c) 2023. StulSoft
 */

package com.stulsoft.poi.examples

import com.typesafe.scalalogging.LazyLogging
import org.apache.poi.ss.usermodel.Workbook
import org.apache.poi.xssf.usermodel.XSSFWorkbook
import scala.jdk.CollectionConverters.*

import java.io.{File, FileInputStream}

object ReadFromSheetEx1 extends LazyLogging {

  private def test1(): Unit = {
    logger.info("==>test1")
    try {
      val fileLocation = "src/test/resources/Test1.xlsx"
      val file = new FileInputStream(new File(fileLocation))
      val workbook = new XSSFWorkbook(file)
      logger.info("Number of sheets: {}", workbook.asScala.size);

      val sheet = workbook.getSheetAt(0).asScala
      logger.info("sheet.size={}", sheet.size)
      for (row <- sheet) {
        logger.info("{} row", row.getRowNum)
        row.getCell(1)
        for (cell <- row.asScala) {
          logger.info("{} column", cell.getColumnIndex)
          logger.info("cell.getCellType: {}", cell.getCellType)
          if cell.getCellType == org.apache.poi.ss.usermodel.CellType.NUMERIC then
            logger.info("numeric value: {}", cell.getNumericCellValue)
            logger.info("numeric value cell style: {}", cell.getCellStyle.toString)
        }
      }

      workbook.close()
    } catch
      case exception: Exception => logger.error(exception.getMessage, exception)
  }

  def main(args: Array[String]): Unit = {
    logger.info("==>main")

    test1()
  }
}
