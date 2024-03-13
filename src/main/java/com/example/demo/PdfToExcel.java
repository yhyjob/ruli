package com.example.demo;

import java.awt.*;
import java.awt.datatransfer.*;
import java.io.IOException;
import java.util.logging.Level;
import java.util.logging.Logger;

class PdfToExcel extends Thread implements ClipboardOwner {
    Clipboard sysClip = Toolkit.getDefaultToolkit().getSystemClipboard();

    public void run() {
        Transferable trans = sysClip.getContents(this);
        regainOwnership(trans);
        System.out.println("Listening to board...");
        while (true) {
        }
    }

    public void lostOwnership(Clipboard c, Transferable t) {
        try {
            sleep(200);
        } catch (Exception e) {
            System.out.println("Exception: " + e);
        }
        Transferable contents = sysClip.getContents(this);
        processContents(contents);
        regainOwnership(contents);
    }

    void processContents(Transferable t) {
        DataFlavor[] flavors = t.getTransferDataFlavors();
        for (int i = flavors.length - 1; i >= 0; i--) {
            try {
                Object o = t.getTransferData(flavors[i]);
//                System.out.println("Flavor " + i + " gives " + o.getClass().getName());
                if (o instanceof String) {
                    System.out.println("String=" + o);

                    break;
                }
            } catch (Exception exp) {
                exp.printStackTrace();
            }
        }
        System.out.println("Processing: ");
    }

    void regainOwnership(Transferable t) {
        sysClip.setContents(t, this);
    }

    public static void main(String[] args) {
//        PdfToExcel b = new PdfToExcel();
//        b.start();

        String initial = "";
        while (true) {
            try {
                Clipboard c = Toolkit.getDefaultToolkit().getSystemClipboard();
                String paste = c.getContents(null).getTransferData(DataFlavor.stringFlavor).toString();
                if (!paste.equals(initial)) {
                    System.out.println(paste);
                    initial = paste;
                }
            } catch (UnsupportedFlavorException | IOException ex) {
                Logger.getLogger(PdfToExcel.class.getName()).log(Level.SEVERE, null, ex);
            }
            try {
                Thread.sleep(400);
            } catch (InterruptedException ex) {
            }
        }
    }
}