package multi.languages.excel.convert;

import java.io.*;


/**
 * @author 岳贺伟
 * @description
 * @data
 */
public class Test {

    public static void main(String[] args) throws IOException {
        ConversionTool conversionTool1 = new ConversionTool("xlsx");
        //conversionTool1.propertiesSwitchExcel("biggie-international","C:\\Users\\18505\\Desktop\\新建文件夹\\");
        conversionTool1.excelSwitchProperties("C:\\Users\\18505\\Desktop\\新建文件夹\\","hug-bixin-biggie-service",1,"C:\\Users\\18505\\Desktop\\新建文件夹\\");
        //conversionTool1.mergeProperties("D:\\IDEA\\bixin\\trade-international\\platform\\hug-order-price-service","C:\\Users\\18505\\Desktop\\新建文件夹");
    }

}
