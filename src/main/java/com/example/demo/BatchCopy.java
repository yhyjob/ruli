package com.example.demo;

import org.apache.commons.collections.CollectionUtils;
import org.apache.log4j.Logger;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.awt.print.PrinterException;
import java.io.FileInputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Collection;

public class BatchCopy {

    static final Logger logger = Logger.getLogger(BatchCopy.class);

    public static void main(String[] args) throws IOException, PrinterException, InterruptedException {
        String path = "D:\\workspaceRuli\\";
        getDate(path, "1");
    }
    public static ArrayList getDate(String path, String packno) throws IOException, PrinterException, InterruptedException {

        path = "D:\\workspaceRuli";
        String datafile = path + "\\202410_9.xlsx";
        XSSFWorkbook wb = new XSSFWorkbook(new FileInputStream(datafile));
        XSSFSheet sheet = wb.getSheetAt(0);
        ArrayList copylist = new ArrayList();
        ArrayList checklist = new ArrayList();
        if (sheet != null) {
            String fundidfirst = "";

            String packnumcektemp = "";
            String descsheetnametemp = "";
            String fundidtemp = "";

            DoublyLinkedList dl = null;
            ArrayList checkSheetlist = null;
            DataFormatter dataFormatter = new DataFormatter();
            for (int rowNum = 1; rowNum <= sheet.getLastRowNum(); rowNum++) {
                //System.out.println(rowNum);
                CopyInfo copyinfo = new CopyInfo();
                XSSFRow hssfRow = sheet.getRow(rowNum);
                Cell packnumcekc = hssfRow.getCell(0);
                String packnumcekcurr = getCellValue(packnumcekc);
                if(("").equals(packnumcekcurr)) break;
                String juesuanqi = dataFormatter.formatCellValue(hssfRow.getCell(1));
                String juesuanyue = dataFormatter.formatCellValue(hssfRow.getCell(2));
                String juesuanri = dataFormatter.formatCellValue(hssfRow.getCell(3));

                String descsheetname = getCellValue(hssfRow.getCell(10));
                String fundid = dataFormatter.formatCellValue((hssfRow.getCell(5)));
                String fundname = getCellValue((hssfRow.getCell(6)));


                if (!packnumcektemp.equals(packnumcekcurr)) {
                    packnumcektemp = packnumcekcurr;
                    descsheetnametemp = descsheetname;
                    fundidtemp = fundid;
                    dl = new DoublyLinkedList();
                    checkSheetlist = new ArrayList();
                    copylist.add(dl);
                    checklist.add(checkSheetlist);
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
                copyinfo.setJuesuanqi(juesuanqi);
                copyinfo.setJuesuanyue(juesuanyue);
                copyinfo.setJuesuanri(juesuanri);
                copyinfo.setFundid(fundidtemp);
                copyinfo.setFundname(fundname);
                copyinfo.setRuliobject(ruliobject);
                copyinfo.setPageno(pageno);
                copyinfo.setDescsheetname(descsheetnametemp);
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
                checkSheetlist.add(copyinfo);
                //logger.info(copyinfo.toString());
            }
        }
        logger.info("开始拷贝");

/*        for (int i = 0; i < copylist.size(); i++) {
            DoublyLinkedList ls = (DoublyLinkedList) copylist.get(i);
            ls.insertTagSet();
            ls.copyListData();
        }*/

        logger.info("开始检查copy和录入数据");
/*        for (int i = 0; i < copylist.size(); i++) {
            DoublyLinkedList ls = (DoublyLinkedList) copylist.get(i);
           // logger.info(ls.getTail().data.getFundid() +"  "+ls.getTail().data.getDescsheetname()+" "+ls.getTail().data.getEndrow());
            ls.checkFileData();

        }*/
        logger.info("开始检查数据是否删除");
        ArrayList sheetall = new ArrayList<>();
        ArrayList sheetchafen = new ArrayList<Integer>();
        for (int i = 1; i < 19; i++) {
            sheetall.add(i);
        }
        for (int i = 0; i < checklist.size(); i++) {
            ArrayList shujulist = (ArrayList) checklist.get(i);
            int q = 0;
            for (int k = 0; k < sheetall.size(); k++) {
                Integer index = (Integer) sheetall.get(k);
                q = k+1;
                CopyInfo copyinfocf = new CopyInfo();
                boolean find = false;
                for (int j = 0; j < shujulist.size(); j++) {
                    CopyInfo copyinfo = (CopyInfo) shujulist.get(j);
                    copyinfocf.setFundid(copyinfo.getFundid());
                    if (index == copyinfo.getSheetindex()) {
                        find = true;
                        break;
                    }
                }
                if (!find) {
                    copyinfocf.setSheetindex(q);
                    sheetchafen.add(copyinfocf);
                }
            }

        }
            DoublyLinkedList.check2(sheetchafen);

        return copylist;
    }

    // 通过 ID 比较，返回 list1 中有而 list2 中没有的对象
/*    public static ArrayList<CopyInfo> getDifference(ArrayList<CopyInfo> list1, ArrayList<CopyInfo> list2) {
        ArrayList<CopyInfo> diff = new ArrayList<>();
        for (CopyInfo p1 : list1) {
            boolean found = false;
            for (CopyInfo p2 : list2) {
                if (p1.getFundid() == p2.getFundid() && p1) {
                    found = true;
                    break;
                }
            }
            if (!found) {
                diff.add(p1);
            }
        }
        return diff;
    }*/

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



