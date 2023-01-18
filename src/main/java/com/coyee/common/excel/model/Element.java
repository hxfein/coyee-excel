package com.coyee.common.excel.model;

/**
 * 所属类别: <br/>
 * 用途: <br/>
 * Author:<a href="mailto:hxfein@126.com">黄飞</a> <br/>
 * Date: 2011-1-9 <br/>
 * Time: 下午09:56:38 <br/>
 * Version: 1.0.2 <br/>
 */
public class Element {
    /**
     * 单元格关键字
     */
    private String key;
    /**
     * 所在列
     */
    private int colIndex;
    /**
     * 所在行
     */
    private int rowIndex;

    public String getKey() {
        return key;
    }

    public void setKey(String key) {
        this.key = key;
    }

    public int getColIndex() {
        return colIndex;
    }

    public void setColIndex(int colIndex) {
        this.colIndex = colIndex;
    }

    public int getRowIndex() {
        return rowIndex;
    }

    public void setRowIndex(int rowIndex) {
        this.rowIndex = rowIndex;
    }

}
