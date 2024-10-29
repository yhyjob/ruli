package com.example.demo;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.*;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;

public class Copysheet {


    public static void copy(String path, String name) throws IOException {
        InputStream is = new FileInputStream(path); // 原始Excel文件
        Workbook workbook = new HSSFWorkbook(is);
        Sheet sheet = workbook.getSheetAt(3); // 假设我们要复制的是第一个工作表
        int col0 = sheet.getColumnWidth(0);
        int col1 = sheet.getColumnWidth(1);
        int col2 = sheet.getColumnWidth(2);


        // 创建新的工作表
        Sheet newSheet = workbook.createSheet(name);
        newSheet.setColumnWidth(0, col0);
        newSheet.setColumnWidth(1, col1);
        newSheet.setColumnWidth(2, col2);

        // 复制所有行
        for (int i = 0; i <= sheet.getLastRowNum(); i++) {
            Row row = sheet.getRow(i);
            if (row == null) {
                continue; // 如果没有数据，跳过
            }
            Row newRow = newSheet.createRow(i);
            newRow.setHeightInPoints(row.getHeightInPoints());
            copyRow(row, newRow);
        }

        // 写入新的Excel文件
        FileOutputStream fileOut = new FileOutputStream(path);
        workbook.write(fileOut);
        fileOut.close();
        workbook.close();
        is.close();
    }

    private static void copyRow(Row srcRow, Row destRow) {
        for (int i = 0; i < srcRow.getLastCellNum(); i++) {
            Cell cell = srcRow.getCell(i);
            if (cell == null) {
                continue; // 如果没有数据，跳过
            }
            Cell newCell = destRow.createCell(i);
            copyCell(cell, newCell);
        }
    }

    private static void copyCell(Cell srcCell, Cell destCell) {
        // 复制单元格样式和值
        CellStyle style = srcCell.getCellStyle();
        destCell.setCellStyle(style);

        switch (srcCell.getCellType()) {
            case STRING:
                destCell.setCellValue(srcCell.getStringCellValue());
                break;
            case NUMERIC:
                destCell.setCellValue(srcCell.getNumericCellValue());
                break;
            case BOOLEAN:
                destCell.setCellValue(srcCell.getBooleanCellValue());
                break;
            case FORMULA:
                destCell.setCellFormula(srcCell.getCellFormula());
                break;
            default:
                break;
        }
    }
}