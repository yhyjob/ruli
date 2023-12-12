package com.example.demo;

import org.apache.pdfbox.pdmodel.PDDocument;
import org.apache.pdfbox.text.PDFTextStripper;

import java.io.File;
import java.io.IOException;

public class Pdfboxreader {
    public static void main(String[] args) throws IOException {
        PDDocument document = PDDocument.load(new File("d:/workspaceRuli/pdf/01311889.pdf"));
        // 2、创建文本提取对象
        PDFTextStripper stripper = new PDFTextStripper();
        // 3、获取指定页面的文本内容
        stripper.setStartPage(18); // 设置起始页面，这里设置成0，就表示读取第一个页面
        stripper.setEndPage(18); // 设置结束页面，这里设置成0，就表示读取第一个页面

        String text = stripper.getText(document);
        System.out.println("获取文本内容: " + text);
        // 4、关闭
        document.close();
    }
}
