package com.example.demo;
import com.spire.pdf.PdfDocument;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import javax.print.attribute.HashPrintRequestAttributeSet;
import javax.print.attribute.PrintRequestAttributeSet;
import javax.print.attribute.standard.PageRanges;
import java.awt.print.PageFormat;
import java.awt.print.Paper;
import java.awt.print.PrinterException;
import java.awt.print.PrinterJob;
import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;

public class BatchPrint {

    public static void main(String[] args) throws IOException, PrinterException {
        String path = "D:\\workspaceRuli\\2回目作業分";


        getDate("", "1");
        // printData("");
    }


    public static void printData(String filename) {
        //Create a PrinterJob object which is initially associated with the default printer
        PrinterJob printerJob = PrinterJob.getPrinterJob();

        //Create a PageFormat object and set it to a default size and orientation
        PageFormat pageFormat = printerJob.defaultPage();

        //Return a copy of the Paper object associated with this PageFormat
        Paper paper = pageFormat.getPaper();

        //Set the imageable area of this Paper
        paper.setImageableArea(0, 0, pageFormat.getWidth(), pageFormat.getHeight());

        //Set the Paper object for this PageFormat
        pageFormat.setPaper(paper);

        //Create a PdfDocument object
        PdfDocument pdf = new PdfDocument();

        //Load a PDF file
        pdf.loadFromFile("D:\\workspaceRuli\\2回目作業分\\01311124.pdf");

        //Call painter to render the pages in the specified format
        printerJob.setPrintable(pdf, pageFormat);

        //Create a PrintRequestAttributed object
        PrintRequestAttributeSet attributeSet = new HashPrintRequestAttributeSet();

        //Set to duplex printing mode
        //attributeSet.add(Sides.TWO_SIDED_SHORT_EDGE);
        attributeSet.add(new PageRanges(1, 7));


        //Execute printing
        try {
            printerJob.print(attributeSet);
        } catch (PrinterException e) {
            e.printStackTrace();
        }
    }


    public static String getDate(String path, String packno) throws IOException, PrinterException {

        path = "D:\\workspaceRuli\\2回目作業分";
        String datafile = path + "\\202404_2.xlsx";
        XSSFWorkbook wb = new XSSFWorkbook(new FileInputStream(datafile));
        XSSFSheet sheet = wb.getSheetAt(0);
        String printfundid = "";
        String filenamepdf = "";
        PrintRequestAttributeSet pagenolist = new HashPrintRequestAttributeSet();

        //Set to duplex printing mode
        //attributeSet.add(Sides.TWO_SIDED_SHORT_EDGE);

        //Create a PrinterJob object which is initially associated with the default printer
        PrinterJob printerJob = PrinterJob.getPrinterJob();

        //Create a PageFormat object and set it to a default size and orientation
        PageFormat pageFormat = printerJob.defaultPage();

        //Return a copy of the Paper object associated with this PageFormat
        Paper paper = pageFormat.getPaper();

        //Set the imageable area of this Paper
        paper.setImageableArea(0, 0, pageFormat.getWidth(), pageFormat.getHeight());

        //Set the Paper object for this PageFormat
        pageFormat.setPaper(paper);

        //Create a PdfDocument object
        PdfDocument pdf = new PdfDocument();

        //Load a PDF file


        //Call painter to render the pages in the specified format
        printerJob.setPrintable(pdf, pageFormat);
        if (sheet != null) {
            for (int rowNum = 1; rowNum <= sheet.getLastRowNum(); rowNum++) {
                //System.out.println(rowNum);
                XSSFRow hssfRow = sheet.getRow(rowNum);
                Cell packnumcek = hssfRow.getCell(0);
                if (packnumcek == null || "".equals(getCellValue(packnumcek))) break;
                Cell fundidcel = hssfRow.getCell(5);
                Cell pageno = hssfRow.getCell(9);
                Cell pagea = hssfRow.getCell(15);
                Cell pageb = hssfRow.getCell(16);

                if ((pageno == null || pageno.getCellType() == CellType.BLANK || "COPY".equals(getCellValue(pageno)) || "".equals(getCellValue(pageno))) &&
                        !getCellValue(pagea).contains("P") && !getCellValue(pageb).contains("P")) {
                    continue;
                }
                if (Double.parseDouble(getCellValue(packnumcek)) == Double.parseDouble(packno)) {
                    if (printfundid.equals("")) {
                        printfundid = getCellValue(fundidcel);
                        filenamepdf = SearchFile(printfundid);
                        pdf.loadFromFile(path + "\\" + filenamepdf);
                    }

                    String range = getCellValue(pageno);
                    if (!"".equals(range)) {
                        if (range.contains(",")) {
                            String[] rangs = range.split(",");
                            pagenolist.add(new PageRanges(Integer.parseInt(rangs[0].trim()) + 1, Integer.parseInt(rangs[1].trim()) + 1));
                            System.out.println("开始打印第" + packno + "包ID为:" + printfundid + "的第" + pageno + "页");
                            //printerJob.print(pagenolist);
                        } else if (range.contains("-")) {
                            String[] rangs = range.split("-");
                            pagenolist.add(new PageRanges(Integer.parseInt(rangs[0].trim()) + 1, Integer.parseInt(rangs[1].trim()) + 1));
                            System.out.println("开始打印第" + packno + "包ID为:" + printfundid + "的第" + pageno + "页");
                            //printerJob.print(pagenolist);
                        } else {
                            pagenolist.add(new PageRanges(Integer.parseInt(range.trim()) + 1));
                            System.out.println("开始打印第" + packno + "包ID为:" + printfundid + "的第" + pageno + "页");
                            //printerJob.print(pagenolist);
                        }
                    }

                    if (getCellValue(pagea).contains("P") && !"COPY".equals(getCellValue(pagea))) {
                        String rangea = getCellValue(pagea).trim().substring(1);
                        pagenolist.add(new PageRanges(Integer.parseInt(rangea.trim()) + 1));
                        System.out.println("开始打印第" + packno + "包ID为:" + printfundid + "保有口数的第" + rangea+ "页");
                        //printerJob.print(pagenolist);
                    }
                    if (getCellValue(pageb).contains("P") && !"COPY".equals(getCellValue(pageb))) {
                        String rangeb = getCellValue(pageb).trim().substring(1);
                        pagenolist.add(new PageRanges(Integer.parseInt(rangeb.trim()) + 1));
                        System.out.println("开始打印第" + packno + "包ID为:" + printfundid + "全体口数的第" + rangeb+ "页");
                        //printerJob.print(pagenolist);
                    }

                }
            }


        }
        //Create a PrintRequestAttributed object
        return path;
    }

    public String getFile(String fundid) {

        return fundid;
    }

    public static String SearchFile(String name) {
        File folder = new File("D:\\workspaceRuli\\2回目作業分");
        File[] listOfFiles = folder.listFiles();

        String targetname = "";
        for (File f : listOfFiles) {
            String fileName = f.getName();
            if (fileName.contains(name) && fileName.contains("pdf")) {
                targetname = fileName;
                break;
            }

        }
        return targetname;
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

        return cellValue;
    }


}
