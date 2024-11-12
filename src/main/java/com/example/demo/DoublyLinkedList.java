package com.example.demo;

import org.apache.log4j.Logger;
import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.DataFormatter;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.text.ParseException;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Date;

public class DoublyLinkedList {

    static final Logger logger = Logger.getLogger(DoublyLinkedList.class);
    private Node<CopyInfo> head;

    public Node<CopyInfo> getTail() {
        return tail;
    }

    public void setTail(Node<CopyInfo> tail) {
        this.tail = tail;
    }

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

    public void copyListData() throws IOException {
        Node<CopyInfo> current = head;
        while (current != null) {
            CopyInfo rowinfo = current.getData();
            if (rowinfo.isCopy()) {
                logger.info(rowinfo);
                copy(rowinfo);
            }
            current = current.next;
        }
    }

    public void insertTagSet() throws IOException {
        Node<CopyInfo> current = tail;
        Boolean isinsert = false;
        while (current != null) {
           String pageno=  current.data.getPageno();
           if("".equals(pageno) || !"COPY".equals(pageno)){
               isinsert = true;
           }
            if (current.data.isCopy()) {
                current.getData().setIsinsert(isinsert);
            }
            current = current.prev;
        }
    }

    public void checkFileData() throws IOException {
        Node<CopyInfo> current = head;
        while (current != null) {
            CopyInfo rowinfo = current.getData();
                check(rowinfo);
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

    public static void copy(CopyInfo rowinfo) throws IOException {
        String fromfilename = "d:/workspaceRuli/copy/" + SearchFile(rowinfo.getSrcfileno());
        String tofilename = SearchFiledes(rowinfo.getFundid());
        if ("".equals(SearchFile(rowinfo.getSrcfileno()))) {
            fromfilename = "d:/workspaceRuli/des/" + SearchFiledes(rowinfo.getSrcfileno());
        }
        logger.info("fromfilename:" + fromfilename);
        logger.info("tofilename:" + tofilename);
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
            if(sourceRow.getCell(0) ==null) break;
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
            if (rowinfo.getIsinsert()) {
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
                        if(sourceRow.getCell(1)==null){
                            logger.info("指示书标注保有口数copy 但是被copy的对象数据为空！");
                        }else {
                            by = sourceRow.getCell(0).getNumericCellValue();
                        }
                    } else if (!rowinfo.isByCopy() && rowinfo.isQtCopy()) {
                        if(sourceRow.getCell(1)==null){
                          logger.info("指示书标注全体口数copy 但是被copy的对象数据为空！");
                        }else {
                            qt = sourceRow.getCell(1).getNumericCellValue();
                        }
                    } else if (rowinfo.isByCopy() && rowinfo.isQtCopy()) {
                        if(sourceRow.getCell(1)==null){
                            logger.info("指示书标注保有口数copy 但是被copy的对象数据为空！");
                        }else {
                            by = sourceRow.getCell(0).getNumericCellValue();
                        }
                        if(sourceRow.getCell(1)==null){
                            logger.info("指示书标注全体口数copy 但是被copy的对象数据为空！");
                        }else {
                            qt = sourceRow.getCell(1).getNumericCellValue();
                        }
                    }
                }
            }
            for (int i = 6; i <= byqutito.getLastRowNum(); i++) {
                HSSFRow sourceRow = byqutito.getRow(i);
                if (rowinfo.getFundtitle().equals(sourceRow.getCell(2).getStringCellValue())) {
                    if (rowinfo.isByCopy() && !rowinfo.isQtCopy()) {
                        if(by!= 0.0){
                            sourceRow.createCell(0).setCellValue(by);
                        }
                    } else if (!rowinfo.isByCopy() && rowinfo.isQtCopy()) {
                        if(qt!= 0.0) {
                            sourceRow.createCell(1).setCellValue(qt);
                        }
                    } else if (rowinfo.isByCopy() && rowinfo.isQtCopy()) {
                        if(by!= 0.0){
                            sourceRow.createCell(0).setCellValue(by);
                        }
                        if(qt!= 0.0) {
                            sourceRow.createCell(1).setCellValue(qt);
                        }
                    }
                }
            }
        }

        rowinfo.setRealcopycount(realcopycount);
        logger.info("需要拷贝的条数："+rowinfo.getCopycount() +"实际拷贝的条数："+realcopycount);
        if (realcopycount != Double.parseDouble(rowinfo.getCopycount())) {
            logger.error(rowinfo);
            logger.error("出错了！！！！！！！！" + "需要拷贝的条数：" + rowinfo.getCopycount() + "实际拷贝的条数：" + realcopycount);
        }

        FileOutputStream fout = new FileOutputStream("d:/workspaceRuli/des/" + tofilename);
        workbookto.write(fout);
        fout.close();
        workbookfrom.close();
        workbookto.close();

    }

    public static void check(CopyInfo rowinfo) throws IOException {
       // logger.info("rowinfo:" + rowinfo.toString());

        String tofilename = SearchFiledes(rowinfo.getFundid());
        POIFSFileSystem fsto = new POIFSFileSystem(new FileInputStream("d:/workspaceRuli/des/" + tofilename));
        HSSFWorkbook workbookto = new HSSFWorkbook(fsto);
        DataFormatter dataFormatter = new DataFormatter();

        if (!"".equals(rowinfo.getJuesuanqi())) {
            HSSFSheet head = workbookto.getSheet("$header");
            //head check
            String juesuanqi = dataFormatter.formatCellValue(head.getRow(2).getCell(1));
            String juesuanyue = dataFormatter.formatCellValue(head.getRow(3).getCell(1));
            String juesuanri = dataFormatter.formatCellValue(head.getRow(4).getCell(1));
            String fundid = dataFormatter.formatCellValue(head.getRow(6).getCell(1));
            String fundname = dataFormatter.formatCellValue(head.getRow(7).getCell(1));

            if (!juesuanqi.equals(rowinfo.getJuesuanqi())) {
                logger.error("决算期不一致：指示书是 " + rowinfo.getJuesuanqi() + " 作业里是 " + juesuanqi);
            }
            if (!juesuanyue.equals(rowinfo.getJuesuanyue())) {
                logger.error("决算期月不一致：指示书是 " + rowinfo.getJuesuanyue() + " 作业里是 " + juesuanyue);
            }
            if ( !isEqueals(juesuanri,rowinfo.getJuesuanri())) {
                logger.error("决算期日不一致：指示书是 " + rowinfo.getJuesuanri() + " 作业里是 " + juesuanri);
            }
            if (!fundid.equals(rowinfo.getFundid())) {
                logger.error("决算基金id不一致：指示书是 " + rowinfo.getFundid() + " 作业里是 " + fundid);
            }
            if (!fundname.equals(rowinfo.getFundname())) {
                logger.error("决算基金名称不一致：指示书是 " + rowinfo.getFundname() + " 作业里是 " + fundname);
            }
        }else{
            logger.error("决算期不为空了!!!!!!");
        }
        if(rowinfo.getSheetindex()==3|| rowinfo.getSheetindex()==4 ){
            // logger.error("作业文件："+tofilename+" sheet 名："+rowinfo.getDescsheetname() +"总行数:"+workbookto.getSheet(rowinfo.getDescsheetname()).getLastRowNum() );
            return;
        }

        if(rowinfo.getSheetindex()==6 || rowinfo.getSheetindex()==7 ){
           // logger.error("作业文件："+tofilename+" sheet 名："+rowinfo.getDescsheetname() +"总行数:"+workbookto.getSheet(rowinfo.getDescsheetname()).getLastRowNum() );
            return;
        }

        //if(rowinfo.getSheetindex()==17 ){

            if (isNumeric(rowinfo.getBaoyou() ) || isNumeric(rowinfo.getQuanti())) {
                boolean nofind = false;
                HSSFSheet byqutifrom = workbookto.getSheetAt(17);
                double by = 0.0;
                double qt = 0.0;
                for (int i = 6; i <= byqutifrom.getLastRowNum(); i++) {
                    HSSFRow sourceRow = byqutifrom.getRow(i);
                    if (rowinfo.getFundtitle().equals(sourceRow.getCell(2).getStringCellValue())) {
                        nofind = true;
                        if (isNumeric(rowinfo.getBaoyou() )  && !isNumeric(rowinfo.getQuanti())) {
                            by = sourceRow.getCell(0).getNumericCellValue();
                            if(by != Double.parseDouble(rowinfo.getBaoyou())){
                                logger.error("作业文件："+tofilename+" sheet 名："+rowinfo.getDescsheetname() +" 保有口数:"+by+"指示书保有口数："+rowinfo.getBaoyou() );
                            }
                        } else if (!isNumeric(rowinfo.getBaoyou() )  && isNumeric(rowinfo.getQuanti())) {
                            nofind = true;
                            qt = sourceRow.getCell(1).getNumericCellValue();
                            if(qt != Double.parseDouble(rowinfo.getQuanti())){
                                logger.error("作业文件："+tofilename+" sheet 名："+rowinfo.getDescsheetname() +" 全体口数:"+qt+"指示书全体口数："+rowinfo.getQuanti() );
                            }
                        } else if (isNumeric(rowinfo.getBaoyou() )  && isNumeric(rowinfo.getQuanti())) {
                            nofind = true;
                            by = sourceRow.getCell(0).getNumericCellValue();
                            qt = sourceRow.getCell(1).getNumericCellValue();
                            if(by != Double.parseDouble(rowinfo.getBaoyou()) && qt != Double.parseDouble(rowinfo.getQuanti())){
                                logger.error("作业文件："+tofilename+" sheet 名："+rowinfo.getDescsheetname() +" 保有口数:"+by+"指示书保有口数："+rowinfo.getBaoyou() );
                                logger.error("作业文件："+tofilename+" sheet 名："+rowinfo.getDescsheetname() +" 全体口数:"+qt+"指示书全体口数："+rowinfo.getQuanti() );
                            }
                        }
                    }
                }
                if(!nofind){
                    logger.error("作业文件："+tofilename+" sheet 名："+rowinfo.getDescsheetname() +"入力对象 :"+rowinfo.getRuliobject()+"保有口数或全体口数没有被录入");
                }
            }

            //return;
       // }


        HSSFSheet sheetto = workbookto.getSheet(rowinfo.getDescsheetname());
        if(rowinfo.getSheetindex()==17){
            int rowcount = sheetto.getLastRowNum()+1-6;
            if(Double.parseDouble(rowinfo.getCopycount() )!= rowcount){
                logger.error("作业文件："+tofilename+" sheet 名："+rowinfo.getDescsheetname() +"入力对象 :"+rowinfo.getRuliobject()+"条数是："+rowinfo.getCopycount()+"实际条目："+rowcount);
            }
            return;
        }

        int realcopycount = 0;
        for (int i = 6; i <= sheetto.getLastRowNum(); i++) {

            HSSFRow sourceRow = sheetto.getRow(i);
            HSSFCell titlenameCell = sourceRow.getCell(rowinfo.getTitleAtColindex());
            String titlename = "";
            if (titlenameCell != null) {
                titlename = sourceRow.getCell(rowinfo.getTitleAtColindex()).getStringCellValue();
            }
            /*else {
                realcopycount = sheetto.getLastRowNum();
                break;
            }*/
            // TODO: 2024/10/29
            String guoname = sourceRow.getCell(0).getStringCellValue();
            Boolean guoqufen = false;
            if (rowinfo.isGuonei() && guoname.contains("日本")) {
                guoqufen = true;
            } else guoqufen = !rowinfo.isGuonei() && !guoname.contains("日本");

            if (rowinfo.getFundtitle().equals(titlename) && guoqufen) {
                realcopycount++;
            }
        }
        boolean b = Double.parseDouble(rowinfo.getCopycount()) != realcopycount;
        if(b){
           logger.error("作业文件："+tofilename+" sheet 名："+rowinfo.getDescsheetname() +"入力对象 :"+rowinfo.getRuliobject()+"条数是："+rowinfo.getCopycount()+"实际条目："+realcopycount);
        }

        workbookto.close();

    }
    public static void check2(ArrayList sheetchafen) throws IOException {

        for (int i = 0; i <sheetchafen.size() ; i++) {

            CopyInfo copyinfo = (CopyInfo) sheetchafen.get(i);
            String tofilename = SearchFiledes(copyinfo.getFundid());
            POIFSFileSystem fsto = new POIFSFileSystem(new FileInputStream("d:/workspaceRuli/des/" + tofilename));
            HSSFWorkbook workbookto = new HSSFWorkbook(fsto);
            logger.info(copyinfo.getSheetindexCheck());
            HSSFSheet sheetto = workbookto.getSheetAt(copyinfo.getSheetindexCheck());
            if(sheetto.getLastRowNum()>5){
                logger.error("作业文件："+tofilename+" sheet 名："+sheetto.getSheetName() +"数据没有删除" );
            }

            workbookto.close();

        }



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

    public static boolean isNumeric(String strNum) {
        try {
            // 尝试将字符串转换为整数
            Integer.parseInt(strNum);
            return true;
        } catch (NumberFormatException e1) {
            // 如果转换为整数失败，尝试转换为双精度浮点数
            try {
                Double.parseDouble(strNum);
                return true;
            } catch (NumberFormatException e2) {
                // 如果两次转换都失败，则字符串不是数值
                return false;
            }
        }
    }
    public static boolean isEqueals(String dateStr1,String dateStr2) {
        // 定义日期格式
        SimpleDateFormat format1 = new SimpleDateFormat("yyyy-MM-dd");
        SimpleDateFormat format2 = new SimpleDateFormat("yyyy/MM/dd");

        try {
            // 解析日期字符串
            Date date1 = format1.parse(dateStr1);
            Date date2 = format2.parse(dateStr2);

            // 比较日期
            if (date1.equals(date2)) {
                return  true;
            } else {
                return false;
            }
        } catch (ParseException e) {
            logger.error("日期格式不正确: " + e.getMessage());
        }
        return false;
    }
}