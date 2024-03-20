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
import java.util.ArrayList;
import java.util.Arrays;
import java.util.Collections;
import java.util.Scanner;
import java.util.concurrent.atomic.AtomicInteger;

class ClipboardListener extends Thread implements ClipboardOwner {
    static Clipboard sysClip = Toolkit.getDefaultToolkit().getSystemClipboard();
    public static final String UTF8_BOM = "\uFEFF";

    private final ArrayList<Integer> threadParam;

    ClipboardListener(ArrayList<Integer> threadParam) throws IOException {
        this.threadParam = threadParam;
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
            //参数取得
            HSSFSheet sheetparam = workbooktarget.getSheetAt(0);
            String colcount = "4.0";
            if (!"".equals(sheetparam.getRow(0).getCell(1).getStringCellValue())) {
                colcount = sheetparam.getRow(0).getCell(1).getStringCellValue();
            }
            String delcol = null;
            if (!"".equals(sheetparam.getRow(1).getCell(1).getStringCellValue())) {
                delcol = sheetparam.getRow(1).getCell(1).getStringCellValue();
            }
            String count = null;
            if (!"".equals(sheetparam.getRow(2).getCell(1).getStringCellValue())) {
                count = sheetparam.getRow(2).getCell(1).getStringCellValue();
            }

            HSSFSheet sheettarg = workbooktarget.getSheetAt(1);
            Scanner scanner = new Scanner(contents);
            scanner.useDelimiter("\n");
            String rows_name = "";
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
                        String name = "";
                        String num = "";
                        if(row.split(" ").length == Integer.parseInt(colcount)){
                             name = row.split(" ")[0];
                             num = row.substring(row.indexOf(" ")+1);
                        }else {
                            if (row.contains("－") && numindex > row.indexOf("－")) {
                                numindex = row.indexOf("－");
                            }
                             name = row.substring(0, numindex);
                             num = row.substring(numindex);
                        }



                        if (!rows_name.equals("")) {
                            name = rows_name + name;
                        }
                        String[] ranksfirst = num.split(" ");
                        ArrayList<String> ranks = new ArrayList<>(Arrays.asList(ranksfirst));

                        rows_name = "";
                        System.out.println(name);
                        //删除处理
                        if (delcol != null) {
                            String[] delcolarg = delcol.split(",");
                            for (int i = 0; i < delcolarg.length; i++) {
                                ranks.remove(Integer.parseInt(delcolarg[i]) - 2);
                            }
                            int count1 = Collections.frequency(ranks, "－");
                            if (count1 == Integer.parseInt(colcount) -delcolarg.length- 1 || ranks.size()<Integer.parseInt(colcount)-delcolarg.length-1) {
                                continue;
                            }
                        }

                        //错误行处理
                        if (ranks.size() > Integer.parseInt(colcount) - 1) {
                            sheettarg.getRow(rownum.intValue()).getCell(0).setCellValue(name + ranks.get(0).trim());
                            ranks.remove(0);
                        } else {
                            sheettarg.getRow(rownum.intValue()).getCell(0).setCellValue(name);
                        }
                        for (int i = 0; i < ranks.size(); i++) {
                            System.out.println(ranks.get(i).trim().trim());
                            sheettarg.getRow(rownum.intValue()).getCell(i + 1).setCellValue(ranks.get(i).trim().trim());
                        }

                    }
                    rownum.incrementAndGet();
                }

            }


            fout = new FileOutputStream(diff_filename);
            workbooktarget.write(fout);
            System.out.println("rownum =" + rownum.intValue());

        } catch (
                Exception e) {
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
        ArrayList<Integer> threadParam = new ArrayList<Integer>();
//        Scanner sc = new Scanner(System.in);
//
//        System.out.println(" Please Enter  共几列:");
//        String colcount = sc.nextLine();  // 读取字符串型输入
//
//        System.out.println(" Please Enter  删除列:");
//        String delcol = sc.nextLine();            // 读取整型输入
//
//        System.out.println(" Please Enter  Height:");
//        float count = sc.nextInt();     // 读取float型输入
//
//        System.out.println("Your Information is as  below:");
//        System.out.println("共几列:" + colcount + "\n" +
//                "删除列:" + delcol + "\n" +
//                "总行数:" + count);
        ClipboardListener b = new ClipboardListener(threadParam);
        //b.itisNotEnough();
        b.start();
    }
}