package com.coyee.common.excel;

import java.io.InputStream;
import java.util.Map;

/**
 * 所属类别: 工具类<br/>
 * 用途: 通过模板Excel文件与数据Excel文件比对，读取出数据Excel中的文件内容<br/>
 * Author:<a href="mailto:hxfein@126.com">黄飞</a> <br/> 
 * Date: 2011-1-22 <br/> 
 * Time: 下午04:28:16 <br/> 
 * Version: 1.0.2 <br/>
 */
public interface AbstractReader {
	/**
	 * 将Excel内容读取为Map
	 * @param templateIns
	 * @param dataIns
	 * @return
	 * @throws Exception
	 */
	Map<String, Object> readExcel(InputStream templateIns,
			InputStream dataIns) throws Exception;
}
