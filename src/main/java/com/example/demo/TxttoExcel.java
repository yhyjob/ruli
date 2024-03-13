package com.example.demo;

import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;

import java.io.*;
import java.util.Date;

public class TxttoExcel {
    public static void main(String[] args) throws IOException {

        String data_filename = "d:/workspaceRuli/pdfdata.txt";
        String diff_filename = "d:/workspaceRuli/target_result.xls";
        POIFSFileSystem fstarget = new POIFSFileSystem(new FileInputStream(diff_filename));
        HSSFWorkbook workbooktarget = new HSSFWorkbook(fstarget);
        HSSFSheet sheettarg = workbooktarget.getSheetAt(0);
        try (BufferedReader br = new BufferedReader(new FileReader(data_filename))) {
            String line;
            int rownum = 0;
            while ((line = br.readLine()) != null) {


                String  str = line; //
                String tempstr = str.replace(", ",",");
                String []  strarg= tempstr.split("\t");
                System.out.println(strarg[0]);
                sheettarg.getRow(rownum).getCell(0).setCellValue(strarg[0].trim());
                System.out.println(strarg[1]);
                String [] numarg = strarg[1].split("\\s{2,}");
                System.out.println(numarg[0]);
                for (int i = 1; i <numarg.length ; i++) {
                    System.out.println(numarg[i] +" length ="+numarg[i].length());
                    sheettarg.getRow(rownum).getCell(i).setCellValue(numarg[i].trim());;

                }
                rownum ++;
            }

        }
        FileOutputStream fout = new FileOutputStream(diff_filename);
        workbooktarget.write(fout);
        workbooktarget.close();

    }

}
