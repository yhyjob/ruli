package com.example.demo;

    import com.spire.pdf.PdfDocument;
import javax.print.PrintService;
import javax.print.attribute.HashPrintRequestAttributeSet;
import javax.print.attribute.PrintRequestAttributeSet;
import javax.print.attribute.standard.PageRanges;
import java.awt.print.PageFormat;
import java.awt.print.Paper;
import java.awt.print.PrinterException;
import java.awt.print.PrinterJob;

public class PrintWithSpecifiedPrinter {

    public static void main(String[] args) throws PrinterException {

        //Create a PrinterJob object which is initially associated with the default printer
        PrinterJob printerJob = PrinterJob.getPrinterJob();

        //Specify printer name
//        PrintService myPrintService = findPrintService("\\\\192.168.1.150\\HP LaserJet P1007");
//        printerJob.setPrintService(myPrintService);

        //Create a PageFormat instance and set it to a default size and orientation
        PageFormat pageFormat = printerJob.defaultPage();

        //Return a copy of the Paper object associated with this PageFormat.
        Paper paper = pageFormat.getPaper();

        //Set the imageable area of this Paper.
       // paper.setImageableArea(0, 0, pageFormat.getWidth(), pageFormat.getHeight());

        //Set the Paper object for this PageFormat.
        pageFormat.setPaper(paper);

        //Create a PdfDocument object
        PdfDocument pdf = new PdfDocument();

        //Load a PDF fileD:\workspaceRuli\5回\5回目作業分\02311118.pdf
        pdf.loadFromFile("D:\\workspaceRuli\\5回\\5回目作業分\\02311118.pdf");

        //Call painter to render the pages in the specified format
        printerJob.setPrintable(pdf, pageFormat);

        //Create a PrintRequestAttributeSet object
        PrintRequestAttributeSet attributeSet = new HashPrintRequestAttributeSet();

        //Set print range
        attributeSet.add(new PageRanges(11,12));

        //Execute printing
        try {
            printerJob.print(attributeSet);
        } catch (PrinterException e) {
            e.printStackTrace();
        }
    }

    //Find print service
    private static PrintService findPrintService(String printerName) {

        PrintService[] printServices = PrinterJob.lookupPrintServices();
        for (PrintService printService : printServices) {
            if (printService.getName().equals(printerName)) {

                System.out.print(printService.getName());
                return printService;
            }
        }
        return null;
    }
}
