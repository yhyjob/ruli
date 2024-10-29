package com.example.demo;

import org.apache.pdfbox.pdmodel.PDDocument;
import org.apache.pdfbox.printing.PDFPageable;

import javax.print.PrintService;
import javax.print.PrintServiceLookup;
import javax.print.attribute.HashPrintRequestAttributeSet;
import javax.print.attribute.PrintRequestAttributeSet;
import javax.print.attribute.standard.PageRanges;
import java.awt.print.PrinterException;
import java.awt.print.PrinterJob;
import java.io.File;
import java.io.IOException;

public class PrintPDF {
    public static void main(String[] args) throws IOException, PrinterException {
        // 确保提供了命令行参数，即PDF文件路径

        // 加载PDF文档
        File pdfFile = new File("D:\\workspaceRuli\\202404_4\\4回目作業分\\02311118.pdf");

        PDDocument document = PDDocument.load(pdfFile);

        // 创建一个打印任务
        PrinterJob job = PrinterJob.getPrinterJob();
        job.setPageable(new PDFPageable(document));

        // 查找并设置打印机
        PrintRequestAttributeSet attributeSet = new HashPrintRequestAttributeSet();
        attributeSet.add(new PageRanges(11,12));
        PrintService defaultPrintService = PrintServiceLookup.lookupDefaultPrintService();
        if (defaultPrintService != null) {
            job.setPrintService(defaultPrintService);
        }

        // 执行打印
        job.print();

        // 关闭PDF文档
        document.close();
    }
}