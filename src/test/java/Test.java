import com.coyee.common.excel.ExcelReader;
import net.sf.json.JSONObject;
import org.apache.commons.io.FileUtils;

import java.io.File;
import java.io.InputStream;
import java.util.Map;

/**
 * @author hxfein
 * @className: Test
 * @description:
 * @date 2023/1/18 11:01
 * @version：1.0
 */
public class Test {

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
}
