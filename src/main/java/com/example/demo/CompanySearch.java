package com.example.demo;

import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;
import org.apache.poi.ss.usermodel.Row;

import java.io.*;
import java.nio.charset.Charset;
import java.util.*;
import java.util.function.Function;
import java.util.stream.Collectors;

public class CompanySearch {
    public static void main(String[] args) throws IOException {
        //new CompanySearch().traverseDirectory("D:\\workspaceRuli\\2023");
        //new CompanySearch().getData("");
        new CompanySearch().userData("");

        // System.out.println(new CompanySearch().isJapanese("山九"));
    }

    HashMap datamap = new HashMap();

    public void traverseDirectory(String directoryPath) throws IOException {
        Deque<File> stack = new ArrayDeque<>();
        File directory = new File(directoryPath);

        stack.push(directory);

        while (!stack.isEmpty()) {
            File currentFile = stack.pop();

            if (currentFile.isDirectory()) {
                File[] files = currentFile.listFiles();
                if (files != null) {
                    for (File file : files) {
                        stack.push(file);
                    }
                }
            } else {
                if (currentFile.getAbsolutePath().indexOf("チェック済み") != -1 && currentFile.getAbsolutePath().indexOf("xls") != -1) {
                    System.out.println(currentFile.getName() + " - " + currentFile.getAbsolutePath());
                    getData(currentFile.getAbsolutePath(), 1);
                    getData(currentFile.getAbsolutePath(), 12);
                }
            }
        }
        Set<String> s1 = datamap.keySet();
        String csvFile = "D:\\workspaceRuli\\data.txt";    //文件名

        try (FileWriter fw = new FileWriter(csvFile)) {

            //开始根据键找值
            for (String key : s1) {
                fw.append(key);
                fw.append('\n');
            }
        } catch (IOException e) {
            e.printStackTrace();
        }


    }

    public void getData(String directoryPath, int sheetnumber) throws IOException {

        POIFSFileSystem fs1 = new POIFSFileSystem(new FileInputStream(directoryPath));
        HSSFWorkbook workbooksrc1 = new HSSFWorkbook(fs1);
        HSSFSheet datasheet = workbooksrc1.getSheetAt(1);

        for (int j = 6; j <= datasheet.getLastRowNum(); j++) {
            Row row = datasheet.getRow(j);
            String name = row.getCell(1).getStringCellValue();
            if (name.length() > 0 && !isJapanese(name.charAt(0))) {
                datamap.put(name, "");
            }
        }
        fs1.close();
    }

    public static void userData(String data) {
        String datafile = "D:\\workspaceRuli\\datafile.txt";    //文件名
        String datastr = "";
        try {

            BufferedReader reader = new BufferedReader( new InputStreamReader(new FileInputStream(datafile), "UTF-8"));
            String line;
            while ((line = reader.readLine()) != null) {
                datastr=datastr+" "+line;
            }
            reader.close();
        } catch (IOException e) {
            e.printStackTrace();
        }
        String ssplitResutlFile = "D:\\workspaceRuli\\splitResut.txt";    //文件名
        String test = HalfFullConversion(datastr.substring(1));
        System.out.println(datastr.substring(1));
        List<String> idsList = new ArrayList<>(Arrays.asList(test.split(" ")));
        //ArrayList database = readFile("");
        List datalast1 = getsuffix1("");
        List datalast = getsuffix("");
        //System.out.println(datalast1.size() +"_"+ datalast.size());
//        List readLast = readLast("");
//            try (FileWriter fw = new FileWriter(csvFile)) {
//                for (int i = 0; i < readLast.size(); i++) {
//                fw.append(datalast.get(i).toString());
//                fw.append('\n');
//                }
//            } catch (IOException e) {
//                throw new RuntimeException(e);
//            }

        //System.out.println("222");
        try (BufferedWriter  fw = new BufferedWriter ( new OutputStreamWriter(new FileOutputStream(ssplitResutlFile), "UTF-8"))){
            String name = "";
            int first = 0;
            for (int i = 0; i < idsList.size(); i++) {
                String temp = idsList.get(i).toString();
                if (first != 0) {
                    name += "　" + temp;
                } else {
                    name += temp;
                }
                String nexttemp = "";
                if (i + 1 < idsList.size()) {
                    nexttemp = idsList.get(i + 1).toString();
                }
                //System.out.println(name);
                //database.contains(name) ||
                //System.out.println(datalast1.contains(temp));
                if (datalast1.contains(temp) || (datalast.contains(temp) && !datalast.contains(nexttemp))) {
                   System.out.println(name);
                    fw.append(name);
                    fw.append('\n');
                    first = 0;
                    name = "";
                } else {
                    first++;
                }
            }
        } catch (IOException e) {
            throw new RuntimeException(e);
        }

    }


    public static String HalfFullConversion(String data) {
        char[] chars = data.toCharArray(); // 将字符串转换为char数组
        StringBuilder output = new StringBuilder();
        for (char c : chars) {
            if ((c >= 'a' && c <= 'z') || (c >= 'A' && c <= 'Z')) {
                int unicodeIndex = (int) (c - '\u0021');
                char fullWidthChar = (char) (unicodeIndex + '\uff01');
                output.append(fullWidthChar);
            } else {
                output.append(c);
            }
        }
        //System.out.println(output.toString());
        return output.toString();
    }

    public static ArrayList readFile(String data) {
        ArrayList database = new ArrayList();
        try {
            BufferedReader reader = new BufferedReader(new FileReader("D:\\workspaceRuli\\data.txt"));
            String line;
            while ((line = reader.readLine()) != null) {
                //System.out.println(line);
                database.add(line);
            }
            reader.close();
        } catch (IOException e) {
            e.printStackTrace();
        }
        return database;
    }

    public static List readLast(String data) {
        List database = new ArrayList();
        HashMap key = new HashMap();
        try {
            BufferedReader reader = new BufferedReader(new FileReader("D:\\workspaceRuli\\data.txt"));
            String line;
            while ((line = reader.readLine()) != null) {
                //System.out.println(line);
                List<String> idsList = new ArrayList<>(Arrays.asList(line.split("　")));
                String[] row = line.split("　");
                int len = line.split("　").length - 1;
                database.add(row[len]);
            }
            reader.close();
        } catch (IOException e) {
            e.printStackTrace();
        }
        return getDuplicateElements(database);
    }

    public static List getsuffix(String data) {
        List database = new ArrayList();
        HashMap key = new HashMap();
        try {
            BufferedReader reader = new BufferedReader(new InputStreamReader(new FileInputStream("D:\\workspaceRuli\\suffix.txt"), "UTF-8") );
            String line;
            while ((line = reader.readLine()) != null) {
                //System.out.println(line);
                database.add(line);
            }
            reader.close();
        } catch (IOException e) {
            e.printStackTrace();
        }
        return database;
    }

    public static List getsuffix1(String data) {
        List database = new ArrayList();
        HashMap key = new HashMap();
        try {
            BufferedReader reader = new BufferedReader(new InputStreamReader(new FileInputStream("D:\\workspaceRuli\\suffix1.txt"), "UTF-8") );
            String line;
            while ((line = reader.readLine()) != null) {
                //System.out.println(line);
                database.add(line);
            }
            reader.close();
        } catch (IOException e) {
            e.printStackTrace();
        }
        return database;
    }

    public static <E> List<E> getDuplicateElements(List<E> list) {
        return list.stream() // list 对应的 Stream
                .collect(Collectors.toMap(e -> e, e -> 1, (a, b) -> a + b)) // 获得元素出现频率的 Map，键为元素，值为元素出现的次数
                .entrySet().stream() // 所有 entry 对应的 Stream
                .filter(entry -> entry.getValue() > 10) // 过滤出元素出现次数大于 1 的 entry
                .map(entry -> entry.getKey()) // 获得 entry 的键（重复元素）对应的 Stream
                .collect(Collectors.toList());  // 转化为 List
    }

    public static boolean isJapanese(char c) {
        int codePoint = Character.codePointAt(Character.toString(c), 0);
        return ((codePoint >= 0x4E00 && codePoint <= 0x9FFF) || (codePoint >= 0x30A1 && codePoint <= 0x30FE));
    }
}
