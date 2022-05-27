package multi.languages.excel.convert;

import java.io.*;


/**
 * @author 岳贺伟
 * @description
 * @data
 */
public class Test {

    public static void main(String[] args) throws IOException {

        /**
         *使用说明：需要指定转换的文件类型 例如：xlsx xls
         *
         *合并操作扫描的resources/i18n（所以你的项目下须有）
         *合并操作mergeProperties 需要输入两个参数（
         *                        参数一：输入服务路径
         *                        参数二：输入输出路径）
         *
         * propertiesSwitchExcel（
         *                        参数一：输入服务名（默认是在本项目的resources下）
         *                        参数二：输入输出路径）
         *
         * excelSwitchProperties（
         *                        参数一：文件路径
         *                        参数二：文件名，
         *                        参数三：语言数量（最多是2，如果再多，需改工具中，产生文件名称）
         *                        参数四：输出路径）
         */

        ConversionTool conversionTool1 = new ConversionTool("xlsx");
        conversionTool1.mergeProperties("D:\\IDEA\\bixin\\trade-international\\platform\\hug-order-price-service","C:\\Users\\18505\\Desktop\\新建文件夹");
        conversionTool1.propertiesSwitchExcel("biggie-international","C:\\Users\\18505\\Desktop\\新建文件夹\\");
        conversionTool1.excelSwitchProperties("C:\\Users\\18505\\Desktop\\新建文件夹\\","hug-bixin-biggie-service",1,"C:\\Users\\18505\\Desktop\\新建文件夹\\");

           }

}
