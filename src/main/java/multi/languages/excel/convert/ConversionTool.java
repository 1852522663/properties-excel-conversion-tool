package multi.languages.excel.convert;

import java.io.BufferedReader;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStreamReader;
import java.io.OutputStreamWriter;
import java.util.*;

import org.apache.log4j.Logger;
import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFCellStyle;
import org.apache.poi.hssf.usermodel.HSSFRichTextString;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.hssf.util.HSSFColor;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.core.io.ClassPathResource;
import org.springframework.core.io.Resource;


/**
 * Properties 与 Excel 文件互转
 * 如果是xls使用HSSFWorkbook；
 * 如果是xlsx，使用XSSFWorkbook
 *
 */
public class ConversionTool {

    public static final Logger log = Logger.getLogger(ConversionTool.class);

    private LinkedHashMap<String, String> propertiesContentToExcelMap = new LinkedHashMap<String, String>();
    //创建一个HashMap，它将存储键和值从xls文件提供
    private LinkedHashMap<String, List<String>> excelContentToPropertiesMap = new LinkedHashMap<String, List<String>>();

    //文件类型
    private String fileType = ".xls";
    //文件中sheet下标
    static int sheetIndexNum = 0;


    public ConversionTool() {

    }

    public ConversionTool(String fileType) {
        this.fileType = "." + fileType;
    }

    /**
     * 功能描述
     *
     * @author 岳贺伟
     * Properties-->Excel
     */
    public void propertiesSwitchExcel(String projectName, String outfileDirPath) {
        //项目包名projectName
        //properties转xls输出文件路径outfileDirPath
        try {

            Resource resource = new ClassPathResource("/" + projectName + "/");

            File propertiesFile = resource.getFile();
            //获取到该路径下所有得properties文件
            String[] files = propertiesFile.list();

            for (int i = 0; i < files.length; i++) {
                String[] split = files[i].split("\\.");
                for (int j = 0; j < split.length; j++) {
                    String name = split[0];
                    //获取properties具体文件路径
                    String inputPropertiesFilePath = propertiesFile.getPath() + "\\" + files[i];
                    //获取父文件路径创建xls文件
                    String outExcelFileNamePath = outfileDirPath + name + fileType;
                    log.info("从properties写入Excel文件开始!");
                    this.propertiesFileConvertToExcel(inputPropertiesFilePath, outExcelFileNamePath, name);
                    log.info("从properties写入Excel文件完成!");

                }
            }
        } catch (IOException e) {
            e.printStackTrace();
        }

    }

    /**
     * 功能描述
     *
     * @author 岳贺伟
     * Excel-->Properties
     */
    public void excelSwitchProperties(String fileDirPath, String fileName, Integer num, String outfileDirPath) {

        if (fileType.equals(".xlsx")) {
            //获取需要转换的xls父文件路径 fileDirPath
            //转换xlsx项目名称xlsName;
            //语言数量num;
            String outExcelFileNamePath1 = fileDirPath + fileName + fileType;

            try {
                log.info("从Excel写入properties文件开始!");

                FileInputStream input = new FileInputStream(new File(outExcelFileNamePath1));
                //使用HSSFWorkbook对象创建工作簿
                HSSFWorkbook workBook = new HSSFWorkbook(input);
                // 获取sheet总数量
                int sheetNum = workBook.getNumberOfSheets();

                for (int i = 0; i < sheetNum; i++) {
                    this.xlsxExcelFileConvertToProperties(workBook, outExcelFileNamePath1, num, outfileDirPath, i);
                }

                log.info("从Excel写入properties文件完成!");
            } catch (FileNotFoundException e) {
                e.printStackTrace();
            } catch (IOException e) {
                e.printStackTrace();
            }
        } else {
            //获取需要转换的xls父文件路径 fileDirPath
            //转换xls项目名称xlsName;
            //语言数量num;
            String outExcelFileNamePath1 = fileDirPath + fileName + fileType;
            System.out.println(outExcelFileNamePath1);
            try {
                log.info("从Excel写入properties文件开始!");

                FileInputStream input = new FileInputStream(new File(outExcelFileNamePath1));
                //使用XSSFWorkbook对象创建工作簿
                XSSFWorkbook workBook = new XSSFWorkbook(input);
                // 获取sheet总数量
                int sheetNum = workBook.getNumberOfSheets();

                for (int i = 0; i < sheetNum; i++) {
                    this.xlsExcelFileConvertToProperties(workBook, outExcelFileNamePath1, num, outfileDirPath, i);
                }

                log.info("从Excel写入properties文件完成!");
            } catch (FileNotFoundException e) {
                e.printStackTrace();
            } catch (IOException e) {
                e.printStackTrace();
            }
        }

    }


    /**
     * 功能描述
     *
     * @author 岳贺伟
     * properties 转换内容转换为Excel
     */

    private void propertiesFileConvertToExcel(String inputPropertiesFilePath, String outExcelFileNamePath, String name) {
        System.setProperty("file.encoding", "UTF-8");
        readPropertiesContent(inputPropertiesFilePath);
        writePropertiesContentToExcelFile(outExcelFileNamePath, name);
    }


    /**
     * 功能描述
     *
     * @author 岳贺伟
     * (xlsx)Excel文件转换为Properties
     */
    private void xlsxExcelFileConvertToProperties(HSSFWorkbook workBook, String inputExcelFileNamePath, Integer num, String outfileDirPath, Integer sheetIndex) {
        ConversionTool conversionTool = new ConversionTool();

        // 通过将xls的位置传递给readExcelFileContent()方法，该方法将把键和值从xls加载到HashMap
        conversionTool.xlsxReadExcelFileContent(workBook, inputExcelFileNamePath, sheetIndex);

        //通过传递属性文件的位置来调用writeExcelToPropertiesFile方法。这个方法将把hashMap中的键和值存储到属性文件中
        conversionTool.writeExcelToPropertiesFile(num, outfileDirPath);


    }

    /**
     * 功能描述
     *
     * @author 岳贺伟
     * (xls)Excel文件转换为Properties
     */
    private void xlsExcelFileConvertToProperties(XSSFWorkbook workBook, String inputExcelFileNamePath, Integer num, String outfileDirPath, Integer sheetIndex) {
        ConversionTool conversionTool = new ConversionTool();


        // 通过将xls的位置传递给readExcelFileContent()方法，该方法将把键和值从xls加载到HashMap
        conversionTool.xlsReadExcelFileContent(workBook, inputExcelFileNamePath, sheetIndex);

        //通过传递属性文件的位置来调用writeExcelToPropertiesFile方法。这个方法将把hashMap中的键和值存储到属性文件中
        conversionTool.writeExcelToPropertiesFile(num, outfileDirPath);


    }

    /**
     * 功能描述
     *
     * @author 岳贺伟
     * 读取 properties文件内容
     */
    private void readPropertiesContent(String propertiesFilePath) {

        // 创建包含属性路径的文件对象
        File propertiesFile = new File(propertiesFilePath);
        // 如果属性文件是一个文件，做下面的事情
        if (propertiesFile.isFile()) {
            try {
                // 创建一个FileInputStream来加载属性文件
                FileInputStream fisProp = new FileInputStream(propertiesFile);
                BufferedReader in = new BufferedReader(new InputStreamReader(fisProp, "UTF8"));

                // 创建Properties对象并加载 通过FileInputStream将属性键和值赋给它
                // 注意事项：默认的Properties
                Properties properties = new OrderedProperties();
                //load方法其实就是逐行读取properties配置文件，分隔成两个字符串key和value
                properties.load(in);

                Enumeration<Object> keysEnum = properties.keys();


                while (keysEnum.hasMoreElements()) {
                    String propKey = (String) keysEnum.nextElement();
                    String propValue = properties.getProperty(propKey);

                    Map<String, String> propItem = new HashMap<String, String>();
                    propItem.put(propKey.trim(), propValue.trim());

                    propertiesContentToExcelMap.put(propKey.trim(), propValue.trim());

                }
                fisProp.close();

            } catch (FileNotFoundException e) {
                e.printStackTrace();
            } catch (IOException e) {
                e.printStackTrace();
            }
        }
    }

    /**
     * 功能描述
     *
     * @author 岳贺伟
     * Properties内容写入Excel之中
     */
    private void writePropertiesContentToExcelFile(String excelPath, String name) {

    /*    Jakarta POI 是一套用于访问微软格式文档的Java API。Jakarta POI有很多组件组成，
    其中有用于操作Excel格式文件的HSSF和用于操作Word的HWPF，目前用于操作Excel的HSSF比较成熟。
    */

        // Workbook workbook = WorkbookFactory.create();
        HSSFWorkbook workBook = new HSSFWorkbook();

        //创建一个名为Properties 的sheet
        HSSFSheet worksheet = workBook.createSheet(name);

        // 在当前sheet中创建第一行
        HSSFRow row = worksheet.createRow((short) 0);

        //设置列头样式
        HSSFCellStyle cellStyle = workBook.createCellStyle();

        cellStyle.setFillForegroundColor(HSSFColor.GOLD.index);
        //设置图案样式
        cellStyle.setFillPattern(HSSFCellStyle.SOLID_FOREGROUND);

        HSSFCell cell1 = row.createCell(0);
        //设置第一行第一列名称
        cell1.setCellValue(new HSSFRichTextString("i18nCode"));
        cell1.setCellStyle(cellStyle);

        HSSFCell cell2 = row.createCell(1);
        cell2.setCellValue(new HSSFRichTextString("中文翻译"));
        cell2.setCellStyle(cellStyle);

        //循环把 Properties文件内容一行一行添加到Excel之中
        for (String s : propertiesContentToExcelMap.keySet()) {

            //在sheet之中每次增加一行
            HSSFRow rowOne = worksheet.createRow(worksheet.getLastRowNum() + 1);
            // 在此行之中创建两列
            HSSFCell cellZero = rowOne.createCell(0);
            HSSFCell cellOne = rowOne.createCell(1);

            //从map和set之中提取 key和value值
            String key;
            key = s;
            String value = propertiesContentToExcelMap.get(key);
            // 把提取的值设置到 Excel之中的列
            cellZero.setCellValue(new HSSFRichTextString(key));
            cellOne.setCellValue(new HSSFRichTextString(value));
        }
        try {
            FileOutputStream fosExcel;
            File fileExcel = new File(excelPath);
            fosExcel = new FileOutputStream(fileExcel);
            workBook.write(fosExcel);
            fosExcel.flush();
            fosExcel.close();
        } catch (Exception e) {
            e.printStackTrace();
        }
    }


    /**
     * 功能描述
     * xlsx类型
     *
     * @author 岳贺伟
     * 读取Excel文件
     */
    public void xlsxReadExcelFileContent(HSSFWorkbook workBook, String fileName, Integer sheetIndex) {

        HSSFCell cell1 = null;
        HSSFCell cell2 = null;

        try {
            // 通过调用获取位置0处的 sheet
            HSSFSheet sheet = workBook.getSheetAt(sheetIndex);
            // 创建 sheet的行迭代器
            Iterator<Row> rowIterator = sheet.rowIterator();

            while (rowIterator.hasNext()) {
                // 通过调用创建对row的引用
                HSSFRow row = (HSSFRow) rowIterator.next();
                Iterator<Cell> cellIterator = row.cellIterator();
                // Iterating over each cell
                cell1 = (HSSFCell) cellIterator.next();
                String key = cell1.getRichStringCellValue().toString();
                List<String> s1 = new ArrayList<>();
                while (cellIterator.hasNext()) {

                    if (!cellIterator.hasNext()) {
                        String value = "";
                        //把key和value放置到 properties Map对象之中
                        s1.add(value);
                    } else {
                        cell2 = (HSSFCell) cellIterator.next();
                        String value = cell2.getRichStringCellValue().toString();

                        s1.add(value);

                    }

                    excelContentToPropertiesMap.put(key, s1);
                }
            }


        } catch (Exception e) {
            log.info("没有发生此类元素异常 ..... ");
            e.printStackTrace();
        }
    }

    /**
     * 功能描述
     * xls类型
     *
     * @author 岳贺伟
     * 读取Excel文件
     */
    public void xlsReadExcelFileContent(XSSFWorkbook workBook, String fileName, Integer sheetIndex) {

        XSSFCell cell1 = null;
        XSSFCell cell2 = null;

        try {
            // 通过调用获取位置0处的 sheet
            XSSFSheet sheet = workBook.getSheetAt(sheetIndex);
            // 创建 sheet的行迭代器
            Iterator<Row> rowIterator = sheet.rowIterator();

            while (rowIterator.hasNext()) {
                // 通过调用创建对row的引用
                XSSFRow row = (XSSFRow) rowIterator.next();
                Iterator<Cell> cellIterator = row.cellIterator();
                // Iterating over each cell
                cell1 = (XSSFCell) cellIterator.next();
                String key = cell1.getRichStringCellValue().toString();
                List<String> s1 = new ArrayList<>();
                while (cellIterator.hasNext()) {

                    if (!cellIterator.hasNext()) {
                        String value = "";
                        //把key和value放置到 properties Map对象之中
                        s1.add(value);
                    } else {
                        cell2 = (XSSFCell) cellIterator.next();
                        String value = cell2.getRichStringCellValue().toString();

                        s1.add(value);

                    }

                    excelContentToPropertiesMap.put(key, s1);
                }
            }


        } catch (Exception e) {
            log.info("没有发生此类元素异常 ..... ");
            e.printStackTrace();
        }
    }

    /**
     * 功能描述
     *
     * @author 岳贺伟
     * Excel写回到Properties文件之中
     */
    public void writeExcelToPropertiesFile(Integer num, String outfileDirPath) {
        //文件名称
        String title = "";

        for (int i = 0; i < num; i++) {
            if (i == 0) {
                title = "messages_zn";
            } else if (i == 1) {
                title = "messages_ar";
            }
            String propertiesPath = outfileDirPath + title + sheetIndexNum + ".properties";
            Properties props = new OrderedProperties();

            //创建一个文件对象，该对象将指向属性文件的位置
            File propertiesFile = new File(propertiesPath);


            try {

                //通过传递上述属性文件创建FileOutputStream 并且设置每次覆盖重写
                FileOutputStream xlsFos = new FileOutputStream(propertiesFile, false);

                // 首先将哈希映射键转换为Set，然后对其进行迭代。
                Iterator<String> mapIterator = excelContentToPropertiesMap.keySet().iterator();

                //遍历迭代器属性
                while (mapIterator.hasNext()) {

                    String key = mapIterator.next().toString();

                    List<String> s2 = excelContentToPropertiesMap.get(key);
                    String value = s2.get(i);

                    if (!key.equals("i18nCode")) {
                        //在上面创建的props对象中设置每个属性key与value
                        props.setProperty(key, value);

                    }
                }

                //最后将属性存储到实属性文件中。
                props.store(new OutputStreamWriter(xlsFos, "utf-8"), null);

            } catch (FileNotFoundException e) {

                e.printStackTrace();

            } catch (IOException e) {

                e.printStackTrace();

            }

        }
        sheetIndexNum++;
    }




    /**
     * 功能描述
     *
     * @author 岳贺伟
     * <p>
     *
     * 合并存放下该项目resources/i18n/的properties
     *
     */
    public void mergeProperties(String OutFilePath) throws IOException {

        Resource resource = new ClassPathResource("/i18n/");
        //存放mes
        Map<String, String> propItem = new HashMap<String, String>();
        //存放err
        Map<String, String> propItem1 = new HashMap<String, String>();
        //写入流
        FileInputStream fisProp = null;
        //读取流mes
        FileOutputStream outputStreamMes = new FileOutputStream(OutFilePath + "\\messages.properties");
        //读取流err
        FileOutputStream outputStreamErr = new FileOutputStream(OutFilePath + "\\errorcode.properties");

        Properties propertiesMes = new OrderedProperties();
        Properties propertiesErr = new OrderedProperties();
        File propertiesFile = resource.getFile();
        //获取到该路径下所有得properties文件
        String[] files = propertiesFile.list();
        for (int i = 0; i < files.length; i++) {

            //如果在i18n下不是messages_zh.properties或messages_zh.properties跳过循环
            if (files[i].equals("messages_zh.properties") || files[i].equals("messages.properties")) {
                log.info("mes开始读取");
                String propertiesFilePath = propertiesFile.getPath() + "\\" + files[i];
                try {
                    // 创建一个FileInputStream来加载属性文件
                    fisProp = new FileInputStream(propertiesFilePath);
                    BufferedReader in = new BufferedReader(new InputStreamReader(fisProp, "UTF8"));
                    //load方法其实就是逐行读取properties配置文件，分隔成两个字符串key和value
                    propertiesMes.load(in);
                    Enumeration<Object> keysEnum = propertiesMes.keys();
                    while (keysEnum.hasMoreElements()) {
                        String propKey = (String) keysEnum.nextElement();
                        String propValue = propertiesMes.getProperty(propKey);
                        propItem.put(propKey.trim(), propValue.trim());
                    }
                } catch (FileNotFoundException e) {
                    e.printStackTrace();
                } catch (IOException e) {
                    e.printStackTrace();
                } finally {
                    fisProp.close();
                }

                log.info("mes读取完成");


            } else if (files[i].equals("errorcode.properties") || files[i].equals("errorcode_zh.properties")) {

                log.info("err开始读取");
                String propertiesFilePath = propertiesFile.getPath() + "\\" + files[i];
                try {
                    // 创建一个FileInputStream来加载属性文件
                    fisProp = new FileInputStream(propertiesFilePath);
                    BufferedReader in = new BufferedReader(new InputStreamReader(fisProp, "UTF8"));
                    //load方法其实就是逐行读取properties配置文件，分隔成两个字符串key和value
                    propertiesErr.load(in);
                    Enumeration<Object> keysEnum = propertiesErr.keys();
                    while (keysEnum.hasMoreElements()) {
                        String propKey = (String) keysEnum.nextElement();
                        String propValue = propertiesErr.getProperty(propKey);
                        propItem1.put(propKey.trim(), propValue.trim());
                    }

                } catch (FileNotFoundException e) {
                    e.printStackTrace();
                } catch (IOException e) {
                    e.printStackTrace();
                } finally {
                    fisProp.close();
                }
                log.info("err读取完成");
            }

        }

        log.info("mes开始写入");
        for (String key : propItem.keySet()) {
            // log.info("key:" + key + " " + "Value:" + propItem.get(key));
            propertiesMes.setProperty(key, propItem.get(key));
        }
        propertiesMes.store(new OutputStreamWriter(outputStreamMes, "UTF8"), null);
        log.info("mes写入完成");
        outputStreamMes.close();

        log.info("err开始写入");
        for (String key : propItem1.keySet()) {
            //log.info("key:" + key + " " + "Value:" + propItem.get(key));
            propertiesErr.setProperty(key, propItem1.get(key));
        }
        propertiesErr.store(new OutputStreamWriter(outputStreamErr, "UTF8"), null);
        log.info("err写入完成");
        outputStreamErr.close();
    }

    /**
     * 功能描述
     *
     * @author 岳贺伟
     * <p>
     * 合并指定 目录scanDir\resources\i18n下的properties
     */
    public void mergeProperties(String scanDir,String OutFilePath) throws IOException {
        //扫描完目录下所有的properties文件，LinkedList放入中
        LinkedList<String> filesPath = new LinkedList<String>();

        File file = new File(scanDir);
        log.info("开始扫描");
        if (file.exists()) {
            //得到所有文件
            File[] listFiles = file.listFiles();
            //System.out.println(Arrays.toString(listFiles));
            for (int i = 0; i < listFiles.length; i++) {
                //遍历文件，切割得到core文件
                String[] split = listFiles[i].getName().split("-");

                if (split[split.length - 1].equals("core")) {

                    //获得core，进入core的resources目录
                    File core = new File(String.valueOf(listFiles[i]) + "\\src\\main\\resources\\i18n");
                    //从resources目录找出需要的文件放入filesPath中
                    String[] list = core.list();
                    for (int j = 0; j < list.length; j++) {
                        if (list[j].equals("messages.properties") || list[j].equals("messages_zh.properties")
                                || list[j].equals("errorcode.properties") || list[j].equals("errorcode_zh.properties")) {
                            String s = core + "\\" + list[j];
                            filesPath.add(s);
                        }

                    }

                }

            }
        }

        log.info("扫描完成");
        //合并
        log.info("开始合并");
        //存放mes
        Map<String, String> propItem = new HashMap<String, String>();
        //存放err
        Map<String, String> propItem1 = new HashMap<String, String>();
        //写入流
        FileInputStream fisProp = null;
        //读取流mes
        FileOutputStream outputStreamMes = new FileOutputStream(OutFilePath + "\\messages.properties");
        //读取流err
        FileOutputStream outputStreamErr = new FileOutputStream(OutFilePath + "\\errorcode.properties");

        Properties propertiesMes = new OrderedProperties();
        Properties propertiesErr = new OrderedProperties();

        //获取到该路径下所有得properties文件
        String[] files = new String[filesPath.size()];
        filesPath.toArray(files);
        //String[] files = propertiesFile.list();
        for (int i = 0; i < files.length; i++) {

            String[] split = files[i].split("\\\\");
            //如果在i18n下不是messages_zh.properties或messages_zh.properties跳过循环
            if (split[split.length-1].equals("messages_zh.properties") || split[split.length-1].equals("messages.properties")) {
                log.info("mes开始读取");

                try {
                    // 创建一个FileInputStream来加载属性文件
                    fisProp = new FileInputStream(files[i]);
                    BufferedReader in = new BufferedReader(new InputStreamReader(fisProp, "UTF8"));
                    //load方法其实就是逐行读取properties配置文件，分隔成两个字符串key和value
                    propertiesMes.load(in);
                    Enumeration<Object> keysEnum = propertiesMes.keys();
                    while (keysEnum.hasMoreElements()) {
                        String propKey = (String) keysEnum.nextElement();
                        String propValue = propertiesMes.getProperty(propKey);
                        propItem.put(propKey.trim(), propValue.trim());
                    }
                } catch (FileNotFoundException e) {
                    e.printStackTrace();
                } catch (IOException e) {
                    e.printStackTrace();
                } finally {
                    fisProp.close();
                }

                log.info("mes读取完成");


            } else if (split[split.length-1].equals("errorcode.properties") || split[split.length-1].equals("errorcode_zh.properties")) {

                log.info("err开始读取");

                try {
                    // 创建一个FileInputStream来加载属性文件
                    fisProp = new FileInputStream(files[i]);
                    BufferedReader in = new BufferedReader(new InputStreamReader(fisProp, "UTF8"));
                    //load方法其实就是逐行读取properties配置文件，分隔成两个字符串key和value
                    propertiesErr.load(in);
                    Enumeration<Object> keysEnum = propertiesErr.keys();
                    while (keysEnum.hasMoreElements()) {
                        String propKey = (String) keysEnum.nextElement();
                        String propValue = propertiesErr.getProperty(propKey);
                        propItem1.put(propKey.trim(), propValue.trim());
                    }

                } catch (FileNotFoundException e) {
                    e.printStackTrace();
                } catch (IOException e) {
                    e.printStackTrace();
                } finally {
                    fisProp.close();
                }
                log.info("err读取完成");
            }

        }

        log.info("mes开始写入");
        for (String key : propItem.keySet()) {
            // log.info("key:" + key + " " + "Value:" + propItem.get(key));
            propertiesMes.setProperty(key, propItem.get(key));
        }
        propertiesMes.store(new OutputStreamWriter(outputStreamMes, "UTF8"), null);
        log.info("mes写入完成");
        outputStreamMes.close();

        log.info("err开始写入");
        for (String key : propItem1.keySet()) {
            //log.info("key:" + key + " " + "Value:" + propItem.get(key));
            propertiesErr.setProperty(key, propItem1.get(key));
        }
        propertiesErr.store(new OutputStreamWriter(outputStreamErr, "UTF8"), null);
        log.info("err写入完成");
        outputStreamErr.close();

        log.info("合并完成");
    }


}
