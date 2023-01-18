package com.coyee.common.excel.model;


/**
 * 所属类别: <br/>
 * 用途: <br/>
 * Author:<a href="mailto:hxfein@126.com">黄飞</a> <br/>
 * Date: 2011-1-9 <br/>
 * Time: 下午09:58:24 <br/>
 * Version: 1.0.2 <br/>
 */
public class DynamicListElement extends Element {
    /**
     * 单元格关键字
     */
    private String colKey;
    /***
     * 停止标识
     */
    private String stopFlag;

    public String getColKey() {
        return colKey;
    }

    public void setColKey(String colKey) {
        this.colKey = colKey;
    }

    public String getStopFlag() {
        return stopFlag;
    }

    public void setStopFlag(String stopFlag) {
        this.stopFlag = stopFlag;
    }
}
