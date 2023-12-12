package com.example.demo;

import com.itextpdf.text.pdf.PdfReader;
import com.itextpdf.text.pdf.parser.PdfTextExtractor;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;

public class Pdfdata {
    public static void main(String[] args) {
        System.out.println("hotfix");
        System.out.println(readPdfByPage("d:/workspaceRuli/5331121A.pdf",21,23));
    }

    public static String readPdfByPage(String fileName, int from, int end) {
        String result = "";
        File file = new File(fileName);
        FileInputStream in = null;
        try {
            in = new FileInputStream(fileName);
            // 新建一个PDF解析器对象
            PdfReader reader = new PdfReader(fileName);
            reader.setAppendable(false);
            // 对PDF文件进行解析，获取PDF文档页码
            int size = reader.getNumberOfPages();
            for (int i = from; i <= end && i < size; i++) {
                //一页页读取PDF文本
                String pageStr = PdfTextExtractor.getTextFromPage(reader, i);
                result = result + pageStr + "\n" + "PDF解析第" + i + "页\n";
            }
            reader.close();
        } catch (Exception e) {
            System.out.println("读取PDF文件" + file.getAbsolutePath() + "生失败！" + e);
            e.printStackTrace();
        } finally {
            if (in != null) {
                try {
                    in.close();
                } catch (IOException e1) {
                }
            }
        }
        return result;
    }

}
