package com.example.demo;

import org.openqa.selenium.By;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.chrome.ChromeOptions;

import java.io.IOException;
import java.lang.reflect.Field;
import java.lang.reflect.InvocationTargetException;
import java.lang.reflect.Method;
import java.math.BigDecimal;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.nio.file.StandardCopyOption;
import java.sql.*;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Date;
import java.util.List;
import java.util.Set;

import static java.lang.Thread.sleep;

public class SeleniumWebCrawler {
    public static void main(String[] args) throws IOException {

        //文件zhunb
//        copy();
//        Copysheet.copy("d:/workspaceRuli/目論見書チェック_scliu.xls","79312024");

        // 设置浏览器驱动路径
//        System.setProperty("webdriver.chrome.driver", "./chromedriver.exe");
//
//        // 创建ChromeDriver对象
//        ChromeOptions option = new ChromeOptions();
//        option.addArguments("--remote-allow-origins=*");
//        WebDriver driver = new ChromeDriver(option);
//
//        try {
//            // 打开目标网页
//             driver.get("https://disclosure2.edinet-fsa.go.jp/weee0050.aspx?c3RpPVoxNjAlMkMmb3RkPTAzMCUyQzA0MCUyQyZ0MT0yMDI0MDYyNiZ0Mj0yMDI0MDYyNiZrYm49MiZwPTE=");
//             sleep(7000);
//             //driver.get("https://www.baidu.com");
//            WebElement table =  driver.findElement(new By.ById("GridContainerTbl"));
//            TableUtil tabledata =  new TableUtil(table);
//            int  rowcount = tabledata.getRowCount();
//            int  colcount  =  tabledata.getColumnCount();
//            for (int i = 2; i <rowcount ; i++) {
//                WebElement td =  tabledata.getWebElementInCell(i,2,By.className("form-control-static")).findElement(By.tagName("span"));
//                String  winhandleBefore =  driver.getWindowHandle();
//                System.out.println(winhandleBefore);
//                td.click();
//                sleep(2000);
//                //System.out.println(td.toString());
//
//               Set<String> winhandles = driver.getWindowHandles();
//                for ( String winhandleAfter:winhandles ) {
//                    if ( winhandleBefore.equals(winhandleAfter)) {
//                        continue;
//                    }
//                    driver.switchTo().window(winhandleAfter);
//                }
//                String  currhand1 =  driver.getWindowHandle();
//                System.out.println(currhand1);
//                System.out.println(driver.getCurrentUrl());
//
//
//                // 菜单选项的取的。
//                driver.switchTo().frame("frame_mokuji");
//                WebElement menus = driver.findElement(By.className("indent-2"));
//                menus.click();
//                sleep(1000);
//                driver.switchTo().defaultContent();
//                sleep(1000);
//                driver.switchTo().frame("frame_honbun");
//                WebElement fundname  = driver.findElement(By.name("jpsps_cor:FundNameTextBlock"));
//
//
//                System.out.println("=========="+fundname.getText());
//            }
//        } catch (Exception e) {
//
//            e.printStackTrace();
//        } finally {
//            // 关闭浏览器驱动
//            driver.quit();
//        }

        // 取的id
       //String[] resultArray = queryArray(aSQL1);
        //System.out.println("==="+sql5);



    }


    public static void copy() {
        SimpleDateFormat dateFormat = new SimpleDateFormat("yyyy-MM-dd");
        Path sourcePath = Paths.get("d:/workspaceRuli/目論見書チェック.xls");
        Path destinationPath = Paths.get("d:/workspaceRuli/目論見書チェック_scliu.xls");

        try {
            Files.copy(sourcePath, destinationPath, StandardCopyOption.REPLACE_EXISTING);
            System.out.println("File copied successfully.");
        } catch (IOException e) {
            System.out.println("An error occurred while copying the file.");
            e.printStackTrace();
        }
    }

}
