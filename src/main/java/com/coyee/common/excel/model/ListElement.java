package com.coyee.common.excel.model;


/**
 * 所属类别: <br/>
 * 用途: <br/>
 * Author:<a href="mailto:hxfein@126.com">黄飞</a> <br/>
 * Date: 2011-1-9 <br/>
 * Time: 下午09:58:24 <br/>
 * Version: 1.0.2 <br/>
 */
public class ListElement extends Element {
    /**
     * 单元格关键字
     */
    private String[] colKeys;
    /***
     * 停止标识
     */
    private String stopFlag;


    public String[] getColKeys() {
        return colKeys;
    }

    public void setColKeys(String[] colKeys) {
        this.colKeys = colKeys;
    }

    public String getStopFlag() {
        return stopFlag;
    }

    public void setStopFlag(String stopFlag) {
        this.stopFlag = stopFlag;
    }
}
