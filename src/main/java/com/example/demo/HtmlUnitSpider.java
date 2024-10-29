package com.example.demo;


import org.htmlunit.WebClient;
import org.htmlunit.html.HtmlPage;

public class HtmlUnitSpider {
    public static void main(String[] args) {
        try (final WebClient webClient = new WebClient()) {
            // 启用JavaScript支持
            webClient.getOptions().setJavaScriptEnabled(true);

            // 禁用Css渲染
            webClient.getOptions().setCssEnabled(false);
            webClient.setJavaScriptTimeout(90000);
            webClient.getOptions().setUseInsecureSSL(true);//接受任何主机连接 无论是否有有效证书
            webClient.getOptions().setJavaScriptEnabled(true);//设置支持javascript脚本
            webClient.getOptions().setCssEnabled(false);//禁用css支持
//            webClient.getOptions().set
            webClient.getOptions().setThrowExceptionOnScriptError(false);//js运行错误时不抛出异常
            webClient.getOptions().setTimeout(100000);//设置连接超时时间
            webClient.getOptions().setDoNotTrackEnabled(false);
            // 获取网页
            HtmlPage page = webClient.getPage("https://disclosure2.edinet-fsa.go.jp/weee0050.aspx?c3RpPVoxNjAlMkMmb3RkPTAzMCUyQzA0MCUyQyZ0MT0yMDI0MDYyNiZ0Mj0yMDI0MDYyNiZrYm49MiZwPTE=");

            // 打印网页内容
            System.out.println(page.asXml());
        } catch (Exception e) {
            e.printStackTrace();
        }
    }
}
