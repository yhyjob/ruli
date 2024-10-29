package com.example.demo;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;
import org.apache.poi.ss.usermodel.Cell;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;

public class DoublyLinkedList {
    private Node<CopyInfo> head;
    private Node<CopyInfo> tail;
    private int size;

    public void add(CopyInfo data) {
        Node<CopyInfo> newNode = new Node<>(data);
        if (head == null) {
            head = newNode;
            tail = newNode;
        } else {
            tail.next = newNode;
            newNode.prev = tail;
            tail = newNode;
        }
        size++;
    }

    public void printList() throws IOException {
        Node<CopyInfo> current = head;
        while (current != null) {
            CopyInfo rowinfo = current.getData();
            if (rowinfo.isCopy()) {
                Boolean isinsert = false;
                if (current.next != null) {
                    isinsert = !current.next.data.isCopy();
                }
                copy(rowinfo, isinsert);
            }
            current = current.next;
        }
    }

    public int endrownum() {
        int rownum = 6;
        Node<CopyInfo> current = head;
        while (current != null) {
            //System.out.println(current.data.getCopycount());
            if (!"".equals(current.data.getCopycount())) {
                rownum = (int) (rownum + Double.parseDouble(current.data.getCopycount()));
            }
            current = current.next;
        }
        return rownum;
    }

    public void printReverse() {
        Node current = tail;
        while (current != null) {
            System.out.print(current.data + " ");
            current = current.prev;
        }
        System.out.println();
    }


    public static String SearchFile(String name) {
        File folder = new File("D:\\workspaceRuli\\copy\\");
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

    public static String SearchFiledes(String name) {
        File folder = new File("D:\\workspaceRuli\\des\\");
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

    public static void copy(CopyInfo rowinfo, Boolean isinsert) throws IOException {
        String fromfilename = "d:/workspaceRuli/copy/" + SearchFile(rowinfo.getSrcfileno());
        String tofilename = SearchFiledes(rowinfo.getFundid());
        if ("".equals(SearchFile(rowinfo.getSrcfileno()))) {
            fromfilename = "d:/workspaceRuli/des/" + SearchFiledes(rowinfo.getSrcfileno());
        }
        System.out.println("fromfilename:" + fromfilename);
        System.out.println("tofilename:" + tofilename);
        POIFSFileSystem fsfrom = new POIFSFileSystem(new FileInputStream(fromfilename));
        HSSFWorkbook workbookfrom = new HSSFWorkbook(fsfrom);
        HSSFSheet sheetfrom = workbookfrom.getSheet(rowinfo.getDescsheetname());

        POIFSFileSystem fsto = new POIFSFileSystem(new FileInputStream("d:/workspaceRuli/des/" + tofilename));
        HSSFWorkbook workbookto = new HSSFWorkbook(fsto);
        HSSFSheet sheetto = workbookto.getSheet(rowinfo.getDescsheetname());
        int realcopycount = 0;
        int Startrownm = rowinfo.getBeginrow();
        for (int i = 6; i <= sheetfrom.getLastRowNum(); i++) {
            HSSFRow sourceRow = sheetfrom.getRow(i);
            HSSFCell titlenameCell = sourceRow.getCell(rowinfo.getTitleAtColindex());
            String titlename = "";
            if (titlenameCell != null) {
                titlename = sourceRow.getCell(rowinfo.getTitleAtColindex()).getStringCellValue();
            }
            // TODO: 2024/10/29
            String guoname = sourceRow.getCell(0).getStringCellValue();
            Boolean guoqufen = false;
            if (rowinfo.isGuonei() && !guoname.contains("日本")) {
                guoqufen = true;
            } else guoqufen = !rowinfo.isGuonei() && guoname.contains("日本");

            if (!rowinfo.getFundtitle().equals(titlename) || guoqufen) {
                continue;
            }
            realcopycount++;
            if (isinsert) {
                sheetto.shiftRows(Startrownm, sheetto.getLastRowNum(), 1);
            }
            HSSFRow destRow = sheetto.createRow(Startrownm);
            Startrownm++;
            if (sourceRow != null) {
                for (int j = 0; j < sourceRow.getLastCellNum(); j++) {
                    HSSFCell sourceCell = sourceRow.getCell(j);
                    HSSFCell destCell = destRow.createCell(j);
                    copyCell(sourceCell, destCell);
                }
            }
        }

        if (rowinfo.isByCopy() || rowinfo.isQtCopy()) {
            HSSFSheet byqutifrom = workbookfrom.getSheetAt(17);
            HSSFSheet byqutito = workbookto.getSheetAt(17);
            double by = 0.0;
            double qt = 0.0;
            for (int i = 6; i <= byqutifrom.getLastRowNum(); i++) {
                HSSFRow sourceRow = byqutifrom.getRow(i);
                if (rowinfo.getFundtitle().equals(sourceRow.getCell(2).getStringCellValue())) {
                    if (rowinfo.isByCopy() && !rowinfo.isQtCopy()) {
                        by = sourceRow.getCell(0).getNumericCellValue();
                    } else if (!rowinfo.isByCopy() && rowinfo.isQtCopy()) {
                        qt = sourceRow.getCell(1).getNumericCellValue();
                    } else if (rowinfo.isByCopy() && rowinfo.isQtCopy()) {
                        by = sourceRow.getCell(0).getNumericCellValue();
                        qt = sourceRow.getCell(1).getNumericCellValue();
                    }


                }
            }
            for (int i = 6; i <= byqutito.getLastRowNum(); i++) {
                HSSFRow sourceRow = byqutito.getRow(i);
                if (rowinfo.getFundtitle().equals(sourceRow.getCell(2).getStringCellValue())) {
                    if (rowinfo.isByCopy() && !rowinfo.isQtCopy()) {
                        sourceRow.createCell(0).setCellValue(by);
                    } else if (!rowinfo.isByCopy() && rowinfo.isQtCopy()) {
                        sourceRow.createCell(1).setCellValue(qt);
                    } else if (rowinfo.isByCopy() && rowinfo.isQtCopy()) {
                        sourceRow.createCell(0).setCellValue(by);
                        sourceRow.createCell(1).setCellValue(qt);
                    }
                }
            }
        }

        rowinfo.setRealcopycount(realcopycount);
        //System.out.println("需要拷贝的条数："+rowinfo.getCopycount() +"实际拷贝的条数："+realcopycount);
        if (realcopycount != Double.parseDouble(rowinfo.getCopycount())) {
            System.out.println(rowinfo);
            System.err.println("出错了！！！！！！！！" + "需要拷贝的条数：" + rowinfo.getCopycount() + "实际拷贝的条数：" + realcopycount);
        }

        FileOutputStream fout = new FileOutputStream("d:/workspaceRuli/des/" + tofilename);
        workbookto.write(fout);
        fout.close();
        workbookfrom.close();
        workbookto.close();

    }

    private static void copyCell(Cell sourceCell, Cell destinationCell) {
        if (sourceCell != null) {
            // 复制单元格的值
            switch (sourceCell.getCellType()) {
                case STRING:
                    destinationCell.setCellValue(sourceCell.getStringCellValue());
                    break;
                case NUMERIC:
                    destinationCell.setCellValue(sourceCell.getNumericCellValue());
                    break;
                case BOOLEAN:
                    destinationCell.setCellValue(sourceCell.getBooleanCellValue());
                    break;
                case FORMULA:
                    destinationCell.setCellFormula(sourceCell.getCellFormula());
                    break;
                default:
                    destinationCell.setCellValue("");
            }

            // 复制单元格的样式
//            CellStyle newCellStyle = destinationCell.getSheet().getWorkbook().createCellStyle();
//            newCellStyle.cloneStyleFrom(sourceCell.getCellStyle());
//            destinationCell.setCellStyle(newCellStyle);
        }
    }
}