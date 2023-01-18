package com.coyee.common.excel.model;


/**
 * 所属类别: <br/> 
 * 用途: <br/> 
 * Author:<a href="mailto:hxfein@126.com">黄飞</a> <br/> 
 * Date: 2011-1-10 <br/> 
 * Time: 下午04:07:57 <br/> 
 * Version: 1.0.2 <br/>
 */
public class MixElement extends Element{
	/**
	 * 左侧字符串
	 */
	private String leftString;
	/**
	 * 左侧字符串
	 */
	private String rightString;

	public String getLeftString() {
		return leftString;
	}
	public void setLeftString(String leftString) {
		this.leftString = leftString;
	}
	public String getRightString() {
		return rightString;
	}
	public void setRightString(String rightString) {
		this.rightString = rightString;
	}

}
