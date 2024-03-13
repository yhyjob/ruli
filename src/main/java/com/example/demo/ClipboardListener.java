package com.example.demo;

import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;

import java.awt.*;
import java.awt.datatransfer.Clipboard;
import java.awt.datatransfer.ClipboardOwner;
import java.awt.datatransfer.DataFlavor;
import java.awt.datatransfer.Transferable;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.Scanner;
import java.util.concurrent.atomic.AtomicInteger;

class ClipboardListener extends Thread implements ClipboardOwner {
    static Clipboard sysClip = Toolkit.getDefaultToolkit().getSystemClipboard();
    public static final String UTF8_BOM = "\uFEFF";

    ClipboardListener() throws IOException {
    }

    AtomicInteger rownum = new AtomicInteger(0);

    public void run() {
        sysClip.setContents(sysClip.getContents(this), this);
        while (true) {
            try {
                sleep(200);
            } catch (InterruptedException e) {
                throw new RuntimeException(e);
            }
        }
    }

    public void lostOwnership(Clipboard clipboard, Transferable contents) {
        try {
            sleep(500);
        } catch (InterruptedException e) {
            throw new RuntimeException(e);
        }
        System.out.println("Clipboard content changed");
        try {
            processClipboardContents(clipboard.getContents(this));
        } catch (IllegalStateException e) {
            e.printStackTrace();
            System.out.println("剪切板异常，正在尝试重新监听，请稍后...");
            boolean flag = true;
            while (flag) {
                try {
                    sysClip.setContents(sysClip.getContents(this), this);
                    flag = false;
                    System.out.println("剪切板已就绪");
                } catch (Exception e1) {
                    try {
                        sleep(500);
                    } catch (InterruptedException ex) {
                        throw new RuntimeException(ex);
                    }
                }
            }


        }
    }

    private void processClipboardContents(Transferable contents) {
        if (contents != null && contents.isDataFlavorSupported(DataFlavor.stringFlavor)) {
            try {
                String text = (String) contents.getTransferData(DataFlavor.stringFlavor);
                //System.out.println(text);
                writetoExcel(text);
                //Clipboard clipboard = Toolkit.getDefaultToolkit().getSystemClipboard();
                Transferable transferable = sysClip.getContents(null);
                sysClip.setContents(transferable, this);

            } catch (Exception e) {
                e.printStackTrace();
                System.out.println("剪切板异常，正在尝试重新监听，请稍后...");
                boolean flag = true;
                while (flag) {
                    try {
                        sysClip.setContents(sysClip.getContents(this), this);
                        flag = false;
                        System.out.println("剪切板已就绪");
                    } catch (Exception e1) {
                        try {
                            sleep(500);
                        } catch (InterruptedException ex) {
                            throw new RuntimeException(ex);
                        }
                    }
                }

            }
        }
        sysClip.setContents(contents, this);
    }

    public synchronized void writetoExcel(String contents) {

        FileOutputStream fout = null;

        HSSFWorkbook workbooktarget = null;
        try {
            String diff_filename = "d:/workspaceRuli/target_result.xls";

            POIFSFileSystem fstarget = new POIFSFileSystem(new FileInputStream(diff_filename));
            workbooktarget = new HSSFWorkbook(fstarget);
            HSSFSheet sheettarg = workbooktarget.getSheetAt(0);
            Scanner scanner = new Scanner(contents);
            scanner.useDelimiter("\n");
            String rows_name = "";
            int firstcolcount =0;
            while (scanner.hasNext()) {
                if (contents.contains("\t")) {
                    String[] ranks = scanner.next().split("\t");
                    for (int i = 0; i < ranks.length; i++) {
                        System.out.println(ranks[i].trim());
                        sheettarg.getRow(rownum.intValue()).getCell(i).setCellValue(ranks[i].trim());
                    }
                    rownum.incrementAndGet();
                } else {
                    String row = scanner.next();
                    int numindex = indexOfFirstDigit(row);
                    if (numindex == -1) {
                        rows_name += row;
                        continue;
                    }
                    if (numindex != -1) {
                        if (row.contains("－") && numindex > row.indexOf("－")) {
                            numindex = row.indexOf("－");
                        }
                        String name = row.substring(0, numindex);
                        String num = row.substring(numindex);
                        String[] ranks = num.split(" ");
                        if (!rows_name.equals("")) {
                            name = rows_name;
                        }
//                        if (firstcolcount==0){
//                            firstcolcount =ranks.length;
//                        }
//                        if(firstcolcount<ranks.length){
//                            sheettarg.getRow(rownum.intValue()).getCell(0).setCellValue(name+ranks[1]);
//                        }
                        sheettarg.getRow(rownum.intValue()).getCell(0).setCellValue(name);
                        rows_name = "";
                        System.out.println(name);

                        for (int i = 0; i < ranks.length; i++) {
                            System.out.println(ranks[i].trim());
//                           if(firstcolcount<ranks.length){
//                               sheettarg.getRow(rownum.intValue()).getCell(i + 1).setCellValue(ranks[i+1].trim());
//                           }else {
                               sheettarg.getRow(rownum.intValue()).getCell(i + 1).setCellValue(ranks[i].trim());
//                           }
                        }
                        firstcolcount = ranks.length;
                        rownum.incrementAndGet();
                    }

                }

            }

            fout = new FileOutputStream(diff_filename);
            workbooktarget.write(fout);
            System.out.println("rownum =" + rownum.intValue());

        } catch (Exception e) {
            e.printStackTrace();

        } finally {
            try {
                if (fout != null) fout.close();
                if (workbooktarget != null) workbooktarget.close();
            } catch (IOException e) {
                throw new RuntimeException(e);
            }

        }
    }

    public static int indexOfFirstDigit(String str) {
        for (int i = 0; i < str.length(); i++) {
            if (Character.isDigit(str.charAt(i))) {
                return i;
            }
        }
        return -1;
    }

    private static boolean isNumeric(String str) {
        try {
            Integer.parseInt(str);
            return true;
        } catch (NumberFormatException e) {
            return false;
        }
    }

    public static void main(String[] args) throws IOException {
        ClipboardListener b = new ClipboardListener();
        //b.itisNotEnough();
        b.start();
    }
}