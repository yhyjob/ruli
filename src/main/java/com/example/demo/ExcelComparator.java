package com.example.demo;

import org.apache.poi.ss.usermodel.*;

import java.io.FileInputStream;
import java.util.Iterator;

public class ExcelComparator {

    public static void main(String[] args) throws Exception {
        Workbook workbook1 = WorkbookFactory.create(new FileInputStream("D:\\workspaceRuli\\コピー ～ 目論見書チェック1.xls"));
        Workbook workbook2 = WorkbookFactory.create(new FileInputStream("D:\\workspaceRuli\\目論見書チェック_20240704_scliu.xls"));

        boolean isEqual = compareWorkbooks(workbook1, workbook2);
        System.out.println("Excel files are " + (isEqual ? "equal" : "not equal"));
    }

    private static boolean compareWorkbooks(Workbook workbook1, Workbook workbook2) {
        if (workbook1.getNumberOfSheets() != workbook2.getNumberOfSheets()) {
            System.out.println("workbook1 Sheets count ="+workbook1.getNumberOfSheets() +"  不等 workbook2 Sheets count ="+workbook2.getNumberOfSheets());
           // return false;
        }

        for (int i = 0; i < workbook1.getNumberOfSheets(); i++) {
            Sheet sheet1 = workbook1.getSheetAt(i);
            Sheet sheet2 = workbook2.getSheetAt(i);
            System.out.println("workbook1 Sheets  ="+sheet1.getSheetName() +"  不等  workbook2 Sheets  ="+sheet2.getSheetName());
            if (!compareSheets(sheet1, sheet2)) {
                return false;
            }
        }

        return true;
    }

    private static boolean compareSheets(Sheet sheet1, Sheet sheet2) {
        if (sheet1.getPhysicalNumberOfRows() != sheet2.getPhysicalNumberOfRows()
               // || sheet1.getPhysicalNumberOfCells() != sheet2.getPhysicalNumberOfCells()
        ) {
            System.out.println("workbook1 getPhysicalNumberOfRows  ="+sheet1.getPhysicalNumberOfRows() +"  不等 workbook2 Sheets count ="+sheet2.getPhysicalNumberOfRows());
            //return false;
        }

        for (int i = 0; i < sheet1.getPhysicalNumberOfRows(); i++) {
            Row row1 = sheet1.getRow(i);
            Row row2 = sheet2.getRow(i);

            if (row1 == null && row2 != null || row1 != null && row2 == null) {
                //return false;
            }

            if (row1 != null && row2 != null) {
                if (row1.getPhysicalNumberOfCells() != row2.getPhysicalNumberOfCells()) {
                    //return false;
                }

                for (int j = 0; j < row1.getPhysicalNumberOfCells(); j++) {
                    Cell cell1 = row1.getCell(j);
                    Cell cell2 = row2.getCell(j);

                    if (!compareCells(cell1, cell2)) {
                        //return false;
                    }
                }
            }
        }

        return true;
    }

    private static boolean compareCells(Cell cell1, Cell cell2) {
        if (cell1 == null && cell2 != null || cell1 != null && cell2 == null) {
            //return false;
        }
        System.out.println();

        if (cell1 != null && cell2 != null) {
            if (cell1.getCellType() != cell2.getCellType()) {
                //return false;
            }

            switch (cell1.getCellType()) {
                case STRING:
                    //System.out.println(cell1.getRow());
                    return cell1.getStringCellValue().equals(cell2.getStringCellValue());
                case NUMERIC:
                    return cell1.getNumericCellValue() == cell2.getNumericCellValue();
                case BOOLEAN:
                    return cell1.getBooleanCellValue() == cell2.getBooleanCellValue();
                case FORMULA:
                    return cell1.getCellFormula().equals(cell2.getCellFormula());
                case BLANK:
                    return true;
                default:
                    //return false;
            }
        }

        return true;
    }
}