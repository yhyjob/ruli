package com.example.demo;


import org.apache.log4j.Logger;
import org.apache.poi.hssf.usermodel.HSSFCellStyle;
import org.apache.poi.hssf.usermodel.HSSFFormulaEvaluator;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.hssf.util.HSSFColor;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.Row;

import java.io.*;
import java.text.DecimalFormat;
import java.text.SimpleDateFormat;
import java.util.Date;

public class DemoApplication {

    private static String copyFileUsingStream(File source, File dest) throws IOException {
        InputStream is = null;
        OutputStream os = null;
        String ret = "";
        try {
            if (!dest.exists()) {
                dest.createNewFile();
            }
            is = new FileInputStream(source);
            os = new FileOutputStream(dest);
            byte[] buffer = new byte[1024];
            int length;
            while ((length = is.read(buffer)) > 0) {
                os.write(buffer, 0, length);
            }
            ret = dest.getPath();
        } finally {
            is.close();
            os.close();
        }
        return ret;
    }

    public static String SearchFile(String name) {
        File folder = new File("d:/workspaceRuli");
        File[] listOfFiles = folder.listFiles();

        String targetname = "";
        for (File f : listOfFiles) {
            String fileName = f.getName();
            if (fileName.contains(name)) {
                targetname = fileName;
                break;
            }

        }
        return targetname;
    }

    public static String SearchChat(String name1, String name2) {
        String rstr = "";
        if (name1.length() != name2.length()) {
            rstr = "长度不一致";
            // return rstr;
        }
        for (int i = 0; i < Math.min(name1.length(), name2.length()); i++) {
            if (name1.charAt(i) != name2.charAt(i)) {
                rstr += "第" + (i + 1) + "位：" + name1.charAt(i) + "__" + name2.charAt(i) + " ";
            }
        }

        return rstr;
    }

    public static String getCellValues(Cell cell) {
        String rstr = "";
        if(cell!=null ){
            if(cell.getCellType().getCode()==0){
                DecimalFormat df = new DecimalFormat("#.######");
                rstr = df.format(cell.getNumericCellValue());
            }else {
                rstr  = cell.getStringCellValue();
            }
        }


        return rstr;
    }



    public static void main(String[] args) throws IOException {
        Logger logger = Logger.getLogger(DemoApplication.class);
        String pathname  = "";
        logger.info("===========start==============");
        try {

            POIFSFileSystem fs = new POIFSFileSystem(new FileInputStream("d:/workspaceRuli/11月15回数据.xls"));
            SimpleDateFormat dateFormat = new SimpleDateFormat("yyyy-MM-dd_HH_mm_ss");
            String diff_filename = "d:/workspaceRuli/diff_result_" + dateFormat.format(new Date()) + ".xls";
            copyFileUsingStream(new File("d:/workspaceRuli/diff_template.xls"), new File(diff_filename));
            POIFSFileSystem fstarget = new POIFSFileSystem(new FileInputStream(diff_filename));
            HSSFWorkbook workbooksrc = new HSSFWorkbook(fs);
            HSSFWorkbook workbooktarget = new HSSFWorkbook(fstarget);

            HSSFSheet sheettarg = workbooktarget.getSheetAt(0);
            HSSFCellStyle cell1Style = workbooktarget.createCellStyle();
            cell1Style.setFillPattern(FillPatternType.SOLID_FOREGROUND);

            int starttarget = 6;
            int target1 = 6;
            int errcount = 0;
            for (int i = 0; i < workbooksrc.getNumberOfSheets(); i++) {
                //System.out.println(workbooksrc.getNameIndex(i);
                starttarget = target1;
                String sheetname = workbooksrc.getSheetAt(i).getSheetName();
                String[] sheetnamearg = sheetname.split("-");
                logger.info("sheet_name is " + sheetname);
                String filename = SearchFile(sheetnamearg[0]);

                if ("".equals(filename)) {
                    logger.debug("没有找到对应的文件");
                    continue;
                }
                logger.info("对应的文件是 " + filename);
                POIFSFileSystem fs1 = new POIFSFileSystem(new FileInputStream("d:/workspaceRuli/" + filename));
                HSSFWorkbook workbooksrc1 = new HSSFWorkbook(fs1);
                HSSFSheet datasheet = workbooksrc1.getSheetAt(Integer.parseInt(sheetnamearg[1]));
                logger.info(sheetname + ": 行数一致抽取：" + (workbooksrc.getSheetAt(i).getLastRowNum() + 1) + " 录入：" + (datasheet.getLastRowNum() - 5));
                //sheettarg.getRow(starttarget).getCell(12).setCellValue(sheetname + ": 行数一致抽取：" + (workbooksrc.getSheetAt(i).getLastRowNum() + 1) + " 录入：" + (datasheet.getLastRowNum() - 5));
                int srccount = 0;
                for (Row row : workbooksrc.getSheetAt(i)) {
                    if (row.getCell(0) != null && !"".equals(row.getCell(0).toString())) {
                        srccount++;
                        //logger.info(row.getCell(0).toString());
                    }
                }
                sheettarg.getRow(starttarget).getCell(12).setCellValue(sheetname + ": 行数一致抽取：" + srccount + " 录入：" + (datasheet.getLastRowNum() - 5));
                if (srccount != datasheet.getLastRowNum() - 5) {
                    logger.info(sheetname + ": 行数不一致抽取：" + srccount + " 录入：" + (datasheet.getLastRowNum() - 5));
                    sheettarg.getRow(starttarget).getCell(13).setCellValue(sheetname + ": 行数不一致抽取：" +srccount + " 录入：" + (datasheet.getLastRowNum() - 5));
         /*           starttarget++;
                    target1++;*/
               //     continue;
                }

                for (Row row : workbooksrc.getSheetAt(i)) {
                    for (int j = 0; j < 4; j++) {
                        // System.out.println(j + "=" + row.getCell(j));
                        String value = "";
                        if (row.getCell(j) != null && !" ".equals(row.getCell(j).toString())) {
                            value = row.getCell(j).toString();
                            if(j!=0 && row.getCell(j).getCellType().getCode()==0){
                                DecimalFormat df = new DecimalFormat("#.######");
                                value = getCellValues(row.getCell(j));
//                                logger.info(value);
                            }

                            HSSFColor  black = (HSSFColor) row.getCell(j).getCellStyle().getFillForegroundColorColor();
                            //logger.info(black.getIndex());
                            if(black.getIndex()!=64 && black.getIndex()!=9){
                                cell1Style.setFillForegroundColor(black.getIndex());
                                sheettarg.getRow(starttarget).getCell(j).setCellStyle(cell1Style);
                            }


                        }
                        //System.out.println(starttarget);
                        sheettarg.getRow(starttarget).getCell(j).setCellValue(value);
                    }
                    starttarget++;
                }
                starttarget = target1;
                // Sh
                for (int j = 6; j <= datasheet.getLastRowNum(); j++) {
                    Row row = datasheet.getRow(j);
                    String value1 = "";
                    String value3 = "";
                    String value5 = "";
                    String value7 = "";
                    if ("1".equals(sheetnamearg[1])) {
                        value1 =getCellValues(row.getCell(1));
                        value3 = getCellValues(row.getCell(3));
                        value5 = getCellValues(row.getCell(5));
                        value7 = getCellValues(row.getCell(7));
                    }
                    if ("12".equals(sheetnamearg[1])) {
                        value1 = getCellValues(row.getCell(1));
                        value3 = getCellValues(row.getCell(2));
                        value5 = getCellValues(row.getCell(4));
                    }

                    sheettarg.getRow(starttarget).getCell(4).setCellValue(value1);
                    sheettarg.getRow(starttarget).getCell(5).setCellValue(value3);
                    sheettarg.getRow(starttarget).getCell(6).setCellValue(value5);
                    sheettarg.getRow(starttarget).getCell(7).setCellValue(value7);
                    if (sheettarg.getRow(starttarget).getCell(0) != null && !sheettarg.getRow(starttarget).getCell(0).toString().equals(value1)) {
                        String before = sheettarg.getRow(starttarget).getCell(0).toString();
                        logger.info("sheetname： " + sheetname + "的第" + (starttarget + 1) + "行数据 " + before + " 与 " + value1 + " 不一致的信息：" + SearchChat(before, value1));
                        sheettarg.getRow(starttarget).getCell(12).setCellValue(SearchChat(before, value1));
                        errcount++;
                    }
                    if (sheettarg.getRow(starttarget).getCell(1) != null && !sheettarg.getRow(starttarget).getCell(1).toString().equals(value3)) {
                        String before = sheettarg.getRow(starttarget).getCell(1).toString();
                        logger.info("sheetname： " + sheetname + "的第" + (starttarget + 1) + "行数据 " + before + " 与 " + value3 + " 不一致的信息：" + SearchChat(before, value3));
                        sheettarg.getRow(starttarget).getCell(12).setCellValue(SearchChat(before, value3));
                        errcount++;
                    }
                    if (sheettarg.getRow(starttarget).getCell(2) != null && !sheettarg.getRow(starttarget).getCell(2).toString().equals(value5)) {
                        String before = sheettarg.getRow(starttarget).getCell(2).toString();
                        logger.info("sheetname： " + sheetname + "的第" + (starttarget + 1) + "行数据 " + before + " 与 " + value5 + " 不一致的信息：" + SearchChat(before, value5));
                        sheettarg.getRow(starttarget).getCell(12).setCellValue(SearchChat(before, value5));
                        errcount++;
                    }
                    if (sheettarg.getRow(starttarget).getCell(3) != null && !sheettarg.getRow(starttarget).getCell(3).toString().equals(value7)) {
                        String before = sheettarg.getRow(starttarget).getCell(3).toString();
                        logger.info("sheetname： " + sheetname + "的第" + (starttarget + 1) + "行数据 " + before + " 与 " + value7 + " 不一致的信息：" + SearchChat(before, value7));
                        sheettarg.getRow(starttarget).getCell(12).setCellValue(SearchChat(before, value7));
                        errcount++;
                    }

                    starttarget++;
                    target1 = starttarget;
                }
                workbooksrc1.close();
            }
            logger.info("错误数：" + errcount);
            HSSFFormulaEvaluator.evaluateAllFormulaCells(workbooktarget);
            FileOutputStream fout = new FileOutputStream(diff_filename);
            workbooktarget.write(fout);
            fout.close();
            workbooksrc.close();
            workbooktarget.close();
        } catch (Exception e) {
            e.printStackTrace();
        }
        logger.info("===========end==============");
    }

}
