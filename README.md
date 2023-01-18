# 复杂excel读取工具类

#### 介绍
提供一种以更简单的读取复杂excel的方法,开发者只需要提根据要读取的excel文件制作模板excel文件,调用ExcelReader.readExcel方法即可完成对跨行/跨列/动态表格/多行动态表格/嵌入式文本内容的读取,可以将几乎所有复杂格式的excel读取为map的形式返回给开发者.
因这段程序初版是在2010年编写的,所以读取部分依赖jxl只能读取.xls格式的文件,后续将支持xlsx的读取.
#### 使用说明
假如要读取一个如下图所示的excel:
![输入图片说明](https://foruda.gitee.com/images/1674012663295980129/4aa6237e_9263307.png "屏幕截图")
1.  编写编写excel读取模板
![输入图片说明](https://foruda.gitee.com/images/1674012729098345310/c142f049_9263307.png "屏幕截图")
2.  调用ExcelReader.readExcel方法
    public static void main(String[] args) throws Exception {
        ExcelReader reader = new ExcelReader();
        InputStream templateIns = Test.class.getResourceAsStream("test-template.xls");
        InputStream dataIns = Test.class.getResourceAsStream("test-data.xls");
        Map<String, Object> dataMap = reader.readExcel(templateIns, dataIns);
        JSONObject json = JSONObject.fromObject(dataMap);
        System.out.println(json.toString());
        templateIns.close();
        dataIns.close();
    }
3.完成读取
![输入图片说明](https://foruda.gitee.com/images/1674012860599102135/08826fd1_9263307.png "屏幕截图")