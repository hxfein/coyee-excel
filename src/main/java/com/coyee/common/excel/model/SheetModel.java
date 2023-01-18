package com.coyee.common.excel.model;

import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;


/**
 * 所属类别: <br/>
 * 用途: <br/>
 * Author:<a href="mailto:hxfein@126.com">黄飞</a> <br/>
 * Date: 2011-1-9 <br/>
 * Time: 下午09:59:46 <br/>
 * Version: 1.0.2 <br/>
 */
public class SheetModel {
    /**
     * 所有单元格
     */
    private List<Element> elements = new ArrayList<Element>();
    /**
     * 单值单元格
     */
    private List<SingleElement> singleElements = new ArrayList<SingleElement>();
    /**
     * 列表单元格
     */
    private List<ListElement> listElements = new ArrayList<ListElement>();

    /**
     * 动态列表单元格
     */
    private List<DynamicListElement> dynamicListElements = new ArrayList<>();
    /**
     * 混值单元格
     */
    private List<MixElement> mixElements = new ArrayList<MixElement>();
    /**
     * 多行列表单元格
     */
    private List<BatchRowListElement> batchRowListElements = new ArrayList<BatchRowListElement>();
    /**
     * 关键字-单元格表
     */
    private Map<String, Element> cells = new HashMap<String, Element>();
    /**
     * 总行数
     */
    private int rows = 0;
    /**
     * 总列数
     */
    private int cols = 0;


    public int getCols() {
        return cols;
    }

    public void setCols(int cols) {
        this.cols = cols;
    }

    public int getRows() {
        return rows;
    }

    public void setRows(int rows) {
        this.rows = rows;
    }

    public List<SingleElement> getSingleElements() {
        return singleElements;
    }

    public void setSingleElements(List<SingleElement> singleElements) {
        this.singleElements = singleElements;
    }

    public List<ListElement> getListElements() {
        return listElements;
    }

    public void setListElements(List<ListElement> listElements) {
        this.listElements = listElements;
    }

    public void addElement(Element el) {
        cells.put(el.getKey(), el);
        elements.add(el);
    }

    public void addSingleElement(SingleElement el) {
        cells.put(el.getKey(), el);
        elements.add(el);
        singleElements.add(el);
    }

    public void addListElement(ListElement el) {
        cells.put(el.getKey(), el);
        elements.add(el);
        listElements.add(el);
    }

    public void addDynamicListElement(DynamicListElement el){
        cells.put(el.getKey(), el);
        elements.add(el);
        dynamicListElements.add(el);
    }

    public void addMixElement(MixElement el) {
        cells.put(el.getKey(), el);
        elements.add(el);
        mixElements.add(el);
    }

    public void addBatchRowListElement(BatchRowListElement el) {
        cells.put(el.getKey(), el);
        elements.add(el);
        batchRowListElements.add(el);
    }

    public Element getElement(String key) {
        return cells.get(key);
    }

    public List<Element> getElements() {
        return elements;
    }

    public void setElements(List<Element> elements) {
        this.elements = elements;
    }

    public List<MixElement> getMixElements() {
        return mixElements;
    }

    public void setMixElements(List<MixElement> mixElements) {
        this.mixElements = mixElements;
    }

    public List<BatchRowListElement> getBatchRowListElements() {
        return batchRowListElements;
    }

    public void setBatchRowListElements(
            List<BatchRowListElement> batchRowListElements) {
        this.batchRowListElements = batchRowListElements;
    }
}
