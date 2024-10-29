package com.example.demo;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.awt.print.PrinterException;
import java.io.FileInputStream;
import java.io.IOException;
import java.util.ArrayList;

public class BatchCopy {

    public static void main(String[] args) throws IOException, PrinterException, InterruptedException {
        String path = "D:\\workspaceRuli\\";
        getDate(path, "1");
    }
    public static ArrayList getDate(String path, String packno) throws IOException, PrinterException, InterruptedException {

        path = "D:\\workspaceRuli";
        String datafile = path + "\\202410_6.xlsx";
        XSSFWorkbook wb = new XSSFWorkbook(new FileInputStream(datafile));
        XSSFSheet sheet = wb.getSheetAt(0);
        ArrayList copylist = new ArrayList();
        if (sheet != null) {
            String fundidfirst = "";

            String packnumcektemp = "";
            String descsheetnametemp = "";
            String fundidtemp = "";

            DoublyLinkedList dl = null;
            DataFormatter dataFormatter = new DataFormatter();
            for (int rowNum = 1; rowNum <= sheet.getLastRowNum(); rowNum++) {
                //System.out.println(rowNum);
                CopyInfo copyinfo = new CopyInfo();
                XSSFRow hssfRow = sheet.getRow(rowNum);
                Cell packnumcekc = hssfRow.getCell(0);
                String packnumcekcurr = getCellValue(packnumcekc);
                String descsheetname = getCellValue(hssfRow.getCell(10));
                String fundid = dataFormatter.formatCellValue((hssfRow.getCell(5)));

                if (!packnumcektemp.equals(packnumcekcurr)) {
                    packnumcektemp = packnumcekcurr;
                    descsheetnametemp = descsheetname;
                    fundidtemp = fundid;
                    dl = new DoublyLinkedList();
                    copylist.add(dl);
                } else {
                    if (!descsheetnametemp.equals(descsheetname) && !descsheetnametemp.equals("")) {
                        descsheetnametemp = descsheetname;
                        dl = new DoublyLinkedList();
                        copylist.add(dl);
                    }
                }
                String ruliobject = getCellValue(hssfRow.getCell(8));
                String pageno = getCellValue(hssfRow.getCell(9));
                String copycount = getCellValue(hssfRow.getCell(11));
                String fundtitle = getCellValue(hssfRow.getCell(12));
                String srcfileno = dataFormatter.formatCellValue(hssfRow.getCell(13));
                String baoyou = getCellValue(hssfRow.getCell(15));
                String quanti = getCellValue(hssfRow.getCell(16));

                copyinfo.setPacknumcek(packnumcekcurr);
                copyinfo.setFundid(fundidtemp);
                copyinfo.setRuliobject(ruliobject);
                copyinfo.setPageno(pageno);
                copyinfo.setDescsheetname(descsheetname);
                copyinfo.setFundtitle(fundtitle);
                copyinfo.setCopycount(copycount);
                copyinfo.setFundtitle(fundtitle);
                copyinfo.setSrcfileno(srcfileno);
                copyinfo.setBaoyou(baoyou);
                copyinfo.setQuanti(quanti);
                dl.add(copyinfo);
                copyinfo.setEndrow(dl.endrownum());
                if (!"".equals(copycount)) {
                    copyinfo.setBeginrow((int) (copyinfo.getEndrow() - Double.parseDouble(copycount)));
                }
                //System.out.println(copyinfo.toString());
            }
        }
        for (int i = 0; i < copylist.size(); i++) {
            DoublyLinkedList ls = (DoublyLinkedList) copylist.get(i);
            ls.printList();


        }

        return copylist;
    }
    public static String getCellValue(Cell cell) {
        String cellValue = "";
        if (cell == null) {
            return cellValue;
        }

        // 判断数据的类型
        switch (cell.getCellType()) {
            case NUMERIC:
                // 数字
                cellValue = String.valueOf(cell.getNumericCellValue());
                break;
            case STRING:
                // 字符串
                cellValue = String.valueOf(cell.getStringCellValue());
                break;
            case BOOLEAN:
                // Boolean
                cellValue = String.valueOf(cell.getBooleanCellValue());
                break;
            case FORMULA:
                // 公式
                cellValue = String.valueOf(cell.getCellFormula());
                break;
            case BLANK:
                // 空值
                cellValue = "";
                break;
            case ERROR:
                // 故障
                cellValue = "非法字符";
                break;
            default:
                cellValue = "未知类型";
                break;
        }

        return cellValue.trim();
    }


    public static void printDate() {

    }
}



