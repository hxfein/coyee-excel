package com.coyee.common.excel;

import com.coyee.common.excel.model.*;
import jxl.*;
import net.sf.json.JSONObject;
import org.apache.commons.io.FileUtils;
import org.apache.commons.lang.math.NumberUtils;
import org.apache.commons.lang.time.DateFormatUtils;

import java.io.File;
import java.io.InputStream;
import java.text.DecimalFormat;
import java.util.*;

/**
 * 所属类别: 工具类<br/>
 * 用途:通过模板Excel文件与数据Excel文件比对，读取出数据Excel中的文件内容, 基于jxl实现 <br/>
 * Author:<a href="mailto:hxfein@126.com">黄飞</a> <br/>
 * Date: 2011-1-10 <br/>
 * Time: 上午09:30:03 <br/>
 * Version: 1.0.2 <br/>
 */
public class ExcelReader implements AbstractReader {
    /**
     * 通过模板Excel文件与数据Excel文件比对，读取出数据Excel中的文件内容
     *
     * @param templateIns
     * @param dataIns
     * @return
     * @throws Exception
     */
    public Map<String, Object> readExcel(InputStream templateIns,
                                         InputStream dataIns) throws Exception {
        return this.readExcel(templateIns, dataIns, 0);
    }

    /**
     * 通过模板Excel文件与数据Excel文件比对，读取出数据Excel中的文件内容
     *
     * @param templateIns
     * @param dataIns
     * @param sheet
     * @return
     * @throws Exception
     */
    public Map<String, Object> readExcel(InputStream templateIns,
                                         InputStream dataIns, int sheet) throws Exception {
        Map<String, Object> data = new HashMap<>();
        Workbook templateBook = Workbook.getWorkbook(templateIns);
        Workbook dataBook = Workbook.getWorkbook(dataIns);
        Sheet templateSheet = templateBook.getSheet(sheet);
        Sheet dataSheet = dataBook.getSheet(sheet);
        // 获取模板信息
        SheetModel sheetModel = getSheetModel(templateSheet, dataSheet);
        List<Element> elements = sheetModel.getElements();
        //按照行的序号从小到大排序，先读取上面的行，再读取下面的行，确保每个元素都被数据顶到了正确的位置
        elements.sort(Comparator.comparingInt(Element::getRowIndex));

        for (Element element : elements) {
            if (element instanceof BatchRowListElement) {
                // 读取多行列表
                BatchRowListElement batchRowListElement = (BatchRowListElement) element;
                this.readBatchRowList(batchRowListElement, dataSheet, sheetModel, data);
            } else if (element instanceof ListElement) {
                // 读取列表
                ListElement listElement = (ListElement) element;
                this.readList(listElement, dataSheet, sheetModel, data);
            }else if (element instanceof DynamicListElement) {
                // 读取列表
                DynamicListElement dynamicListElement = (DynamicListElement) element;
                this.readDynamicList(dynamicListElement, dataSheet, sheetModel, data);
            }else if (element instanceof MixElement) {
                // 读取混合内容
                MixElement mixElement = (MixElement) element;
                this.readMix(mixElement, dataSheet, sheetModel, data);
            } else if (element instanceof SingleElement) {
                // 读取单值内容
                SingleElement singleElement = (SingleElement) element;
                this.readSingle(singleElement, dataSheet, sheetModel, data);
            }
        }
        templateBook.close();
        dataBook.close();
        templateIns.close();
        dataIns.close();
        return data;
    }

    /**
     * 读取列表值单元格
     *
     * @param listElement
     * @param dataSheet
     * @param sheetModel
     * @param data
     */
    private void readList(ListElement listElement, Sheet dataSheet,
                          SheetModel sheetModel, Map<String, Object> data) {
        int cols = dataSheet.getColumns();
        List<Map<String, String>> dataList = new ArrayList<>();
        String key = listElement.getKey();
        String stopFlag = listElement.getStopFlag();
        int colIndex = listElement.getColIndex();
        int rowIndex = listElement.getRowIndex();
        String colKeys[] = listElement.getColKeys();
        int rows = sheetModel.getRows();
        boolean flag = true;

        int stopRowIndex = -1;
        if ("none".equals(stopFlag)) {
            stopRowIndex = rows;
        } else {
            for (int j = 0; j < rows; j++) {
                for (int k = colIndex; k < cols; k++) {
                    Cell cell = dataSheet.getCell(k, j);
                    String val = this.getCellValue(cell);
                    if (val.startsWith(stopFlag)) {
                        stopRowIndex = j;
                        j = rows;
                        break;
                    }
                }
            }
        }
        for (int j = rowIndex; j < stopRowIndex && j < rows; j++) {
            Map<String, String> row = new HashMap<String, String>();
            for (int k = 0; k < colKeys.length; k++) {
                Cell cell = dataSheet.getCell(colIndex + k, j);
                String value = getCellValue(cell);
                row.put(colKeys[k], value);
            }
            if (flag) {
                dataList.add(row);
            }
        }
        // 操作模板信息中行号比此模板定义信息行号大的项，将行号加上用户填写的数据条数，使其重新指向正确的位置
        int rowCount = dataList.size();
        List<Element> elements = sheetModel.getElements();
        int elSize = elements.size();
        for (int j = 0; j < elSize; j++) {
            Element el0 = elements.get(j);
            int rowIndex0 = el0.getRowIndex();
            if (rowIndex0 > rowIndex) {
                el0.setRowIndex(rowIndex0 + rowCount - 1);
            }
        }
        data.put(key, dataList);
    }


    /**
     * 读取不定列数列表值单元格
     *
     * @param listElement
     * @param dataSheet
     * @param sheetModel
     * @param data
     */
    private void readDynamicList(DynamicListElement listElement, Sheet dataSheet,
                          SheetModel sheetModel, Map<String, Object> data) {
        int cols = dataSheet.getColumns();
        List<List<String>> dataList = new ArrayList<>();
        String stopFlag = listElement.getStopFlag();
        int colIndex = listElement.getColIndex();
        int rowIndex = listElement.getRowIndex();
        String colKey = listElement.getColKey();
        int rows = sheetModel.getRows();
        boolean flag = true;

        int stopRowIndex = -1;
        if ("none".equals(stopFlag)) {
            stopRowIndex = rows;
        } else {
            for (int j = 0; j < rows; j++) {
                for (int k = colIndex; k < cols; k++) {
                    Cell cell = dataSheet.getCell(k, j);
                    String val = this.getCellValue(cell);
                    if (val.startsWith(stopFlag)) {
                        stopRowIndex = j;
                        j = rows;
                        break;
                    }
                }
            }
        }
        for (int j = rowIndex; j < stopRowIndex && j < rows; j++) {
            List<String> rowValues=new ArrayList<>();
            for (int k = 0; k < cols; k++) {
                Cell cell = dataSheet.getCell(colIndex + k, j);
                String value = getCellValue(cell);
                rowValues.add(value);
            }
            if (flag) {
                dataList.add(rowValues);
            }
        }
        // 操作模板信息中行号比此模板定义信息行号大的项，将行号加上用户填写的数据条数，使其重新指向正确的位置
        int rowCount = dataList.size();
        List<Element> elements = sheetModel.getElements();
        int elSize = elements.size();
        for (int j = 0; j < elSize; j++) {
            Element el0 = elements.get(j);
            int rowIndex0 = el0.getRowIndex();
            if (rowIndex0 > rowIndex) {
                el0.setRowIndex(rowIndex0 + rowCount - 1);
            }
        }
        data.put(colKey, dataList);
    }


    /**
     * 读取单值单元格
     *
     * @param singleElement
     * @param dataSheet
     * @param sheetModel
     * @param data
     */
    private void readSingle(SingleElement singleElement, Sheet dataSheet,
                            SheetModel sheetModel, Map<String, Object> data) {
        String key = singleElement.getKey();
        int rowIndex = singleElement.getRowIndex();
        int colIndex = singleElement.getColIndex();
        Cell cell = dataSheet.getCell(colIndex, rowIndex);
        String value = getCellValue(cell);
        data.put(key, value);
    }

    /**
     * 读取模板变量定义和模板文本定义在同一个单元格中的情况
     *
     * @param mixElement
     * @param dataSheet
     * @param sheetModel
     * @param data
     */
    private void readMix(MixElement mixElement, Sheet dataSheet,
                         SheetModel sheetModel, Map<String, Object> data) {
        int rowIndex = mixElement.getRowIndex();
        int colIndex = mixElement.getColIndex();
        Cell cell = dataSheet.getCell(colIndex, rowIndex);
        String key = mixElement.getKey();
        String leftString = mixElement.getLeftString();
        String rightString = mixElement.getRightString();
        String value = getCellValue(cell);
        value = value.replaceFirst(leftString, "");
        value = value.replaceFirst(rightString, "");
        data.put(key, value);
    }

    /**
     * 获取模板中定义的信息
     *
     * @param sheet
     * @return
     */
    private SheetModel getSheetModel(Sheet sheet, Sheet dataSheet) {
        SheetModel sheetModel = new SheetModel();
        int rows = sheet.getRows();
        int cols = sheet.getColumns();
        for (int i = 0; i < cols; i++) {
            for (int j = 0; j < rows; j++) {
                Cell cell = sheet.getCell(i, j);
                String content = cell.getContents();
                if (content.startsWith("$S")) {
                    // 单值
                    SingleElement el = getSingleElement(i, j, content);
                    sheetModel.addSingleElement(el);
                } else if (content.startsWith("$BR")) {
                    // 多行列表
                    BatchRowListElement el = getBatchRowList(i, j, content);
                    sheetModel.addBatchRowListElement(el);
                } else if (content.startsWith("$L")) {
                    // 列表
                    ListElement el = getListElement(i, j, content);
                    sheetModel.addListElement(el);
                } else if (content.startsWith("$DL")) {
                    //不定列数列表
                    DynamicListElement el=getDynamicListElement(i,j,content);
                    sheetModel.addDynamicListElement(el);
                } else if (content.indexOf("$M") != -1) {
                    // 混合
                    MixElement el = getMixElement(i, j, content);
                    sheetModel.addMixElement(el);
                } else {
                    // 文本内容
                    Element el = getTextElement(i, j, content);
                    sheetModel.addElement(el);
                }
            }
        }
        sheetModel.setRows(dataSheet.getRows());
        sheetModel.setCols(dataSheet.getColumns());
        return sheetModel;
    }

    /**
     * 封装简单文件单元格
     *
     * @param i
     * @param j
     * @param content
     * @return
     */
    private Element getTextElement(int i, int j, String content) {
        Element el = new Element();
        el.setColIndex(i);
        el.setRowIndex(j);
        el.setKey(content);
        return el;
    }

    /**
     * 封装文本与模板混合的单元格
     *
     * @param i
     * @param j
     * @param content
     * @return
     */
    private MixElement getMixElement(int i, int j, String content) {
        MixElement el = new MixElement();
        int left = content.indexOf("$M<");
        int right = content.indexOf(">", left);
        String key = content.substring(left + 3, right);
        String leftString = content.substring(0, left);
        String rightString = content.substring(right + 1);
        el.setLeftString(leftString);
        el.setRightString(rightString);
        el.setColIndex(i);
        el.setRowIndex(j);
        el.setKey(key);
        return el;
    }

    /**
     * 封装列表类型的单元格
     *
     * @param i
     * @param j
     * @param content
     * @return
     */
    private ListElement getListElement(int i, int j, String content) {
        int left = content.indexOf("<");
        int right = content.indexOf(">");
        int stop = content.indexOf("!");
        String key = content.substring(2, left);
        String names = content.substring(left + 1, right);
        String stopFlag = content.substring(stop + 1);
        ListElement el = new ListElement();
        el.setColIndex(i);
        el.setRowIndex(j);
        el.setColKeys(names.split(","));
        el.setStopFlag(stopFlag);
        el.setKey(key);
        return el;
    }

    /**
     * 封装动态列表类型的单元格
     *
     * @param i
     * @param j
     * @param content
     * @return
     */
    private DynamicListElement getDynamicListElement(int i, int j, String content) {
        int stop = content.indexOf("!");
        String colKey = content.substring(2, stop);
        String stopFlag = content.substring(stop + 1);
        DynamicListElement el = new DynamicListElement();
        el.setColIndex(i);
        el.setRowIndex(j);
        el.setStopFlag(stopFlag);
        el.setColKey(colKey);
        return el;
    }

    /**
     * 获取多行多列的模型
     *
     * @param i
     * @param j
     * @param content
     */
    private BatchRowListElement getBatchRowList(int i, int j, String content) {
        int left = content.indexOf("<");
        int right = content.indexOf(">");
        int stop = content.indexOf("!");
        int tail = content.lastIndexOf("#");
        String key = content.substring(3, left);
        String names = content.substring(left + 1, right);
        String stopFlag = content.substring(stop + 1, tail);
        String startColumn = content.substring(tail + 1);
        int startColumnIndex = NumberUtils.toInt(startColumn);
        String colKeyRows[] = names.split("\\|");
        int skipRows = colKeyRows.length;
        String colKeys[][] = new String[skipRows][];
        for (int k = 0; k < colKeyRows.length; k++) {
            String rowColKeys[] = colKeyRows[k].split(",");
            colKeys[k] = rowColKeys;
        }
        BatchRowListElement el = new BatchRowListElement();
        el.setColIndex(i);
        el.setRowIndex(j);
        el.setColKeys(colKeys);
        el.setStopFlag(stopFlag);
        el.setKey(key);
        el.setStartColIndex(startColumnIndex);
        return el;
    }

    /**
     * 封装单值的单元格
     *
     * @param i
     * @param j
     * @param content
     * @return
     */
    private SingleElement getSingleElement(int i, int j, String content) {
        SingleElement el = new SingleElement();
        String key = content.substring(2);
        el.setColIndex(i);
        el.setRowIndex(j);
        el.setKey(key);
        return el;
    }

    private String getCellValue(Cell cell) {
        if (cell.getType() == CellType.NUMBER) {
            DecimalFormat formator = new DecimalFormat("#0.00");
            return formator.format((((NumberCell) cell).getValue()));
        } else if (cell.getType() == CellType.DATE) {
            Date date = ((DateCell) cell).getDate();
            return DateFormatUtils.format(date, "yyy-MM-dd");
        } else {
            return cell.getContents();
        }
    }

    /**
     * 获取多行列表的数据
     *
     * @param batchRowListElement
     * @param dataSheet
     * @param sheetModel
     * @param data
     */
    private void readBatchRowList(BatchRowListElement batchRowListElement, Sheet dataSheet,
                                  SheetModel sheetModel, Map<String, Object> data) {
        int cols = sheetModel.getCols();
        int rows = sheetModel.getRows();
        List<Map<String, String>> dataList = new ArrayList<>();
        String key = batchRowListElement.getKey();
        String stopFlag = batchRowListElement.getStopFlag();
        int colIndex = batchRowListElement.getColIndex();
        int rowIndex = batchRowListElement.getRowIndex();
        String colKeys[][] = batchRowListElement.getColKeys();
        int skipRow = colKeys.length;
        int startColIndex = batchRowListElement.getStartColIndex();

        // 查找停止行
        int stopRowIndex = -1;
        for (int j = 0; j < rows; j++) {
            for (int k = colIndex; k < cols; k++) {
                Cell cell = dataSheet.getCell(k, j);
                String val = this.getCellValue(cell);
                if (val.startsWith(stopFlag)) {
                    stopRowIndex = j;
                    j = rows;
                    break;
                }
            }
        }
        stopRowIndex = stopRowIndex == -1 ? rows : stopRowIndex;
        // 循环找值
        for (int j = rowIndex; j < stopRowIndex; j += skipRow) {
            Map<String, String> map = new HashMap<>();
            for (int k = 0; k < skipRow; k++) {
                String rowColKeys[] = colKeys[k];
                for (int m = startColIndex, n = 0; m < cols && n < rowColKeys.length; m++, n++) {
                    Cell cell = dataSheet.getCell(m, j + k);
                    String val = getCellValue(cell);
                    String mapKey = rowColKeys[n];
                    map.put(mapKey, val);
                }
            }
            dataList.add(map);
        }
        data.put(key, dataList);


        // 操作模板信息中行号比此模板定义信息行号大的项，将行号加上用户填写的数据条数，使其重新指向正确的位置
        int rowCount = dataList.size();
        List<Element> elements = sheetModel.getElements();
        int elSize = elements.size();
        for (int j = 0; j < elSize; j++) {
            Element el0 = elements.get(j);
            int rowIndex0 = el0.getRowIndex();
            if (rowIndex0 > rowIndex) {
                el0.setRowIndex(rowIndex0 + rowCount * skipRow - 1);
            }
        }

    }
}
