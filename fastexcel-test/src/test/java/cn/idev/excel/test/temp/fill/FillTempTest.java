package cn.idev.excel.test.temp.fill;

import java.io.File;
import java.util.ArrayList;
import java.util.Date;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

import cn.idev.excel.EasyExcel;
import cn.idev.excel.ExcelWriter;
import cn.idev.excel.enums.WriteDirectionEnum;
import cn.idev.excel.test.demo.fill.FillData;
import cn.idev.excel.test.util.TestFileUtil;
import cn.idev.excel.write.metadata.WriteSheet;
import cn.idev.excel.write.metadata.fill.FillConfig;
import cn.idev.excel.write.metadata.fill.FillWrapper;

import org.junit.jupiter.api.Test;

/**
 * Example of filling data into Excel templates.
 *
 * @author Jiaju Zhuang
 * @since 2.1.1
 */

public class FillTempTest {
    /**
     * Simplest example of filling data.
     *
     * @since 2.1.1
     */
    @Test
    public void simpleFill() {
        // Template note: Use {} to represent variables. If the template contains "{", "}" as special characters, use "\{", "\}" instead.
        String templateFileName = "src/test/resources/fill/simple.xlsx";

        // Option 1: Fill using an object
        String fileName = TestFileUtil.getPath() + "simpleFill" + System.currentTimeMillis() + ".xlsx";
        // This will fill the first sheet, and the file stream will be closed automatically.
        FillData fillData = new FillData();
        fillData.setName("Zhang San");
        fillData.setNumber(5.2);
        EasyExcel.write(fileName).withTemplate(templateFileName).sheet().doFill(fillData);

        /*
        // Option 2: Fill using a Map
        fileName = TestFileUtil.getPath() + "simpleFill" + System.currentTimeMillis() + ".xlsx";
        // This will fill the first sheet, and the file stream will be closed automatically.
        Map<String, Object> map = new HashMap<String, Object>();
        map.put("name", "Zhang San");
        map.put("number", 5.2);
        EasyExcel.write(fileName).withTemplate(templateFileName).sheet().doFill(map);
        */
    }

    /**
     * Example of filling a list of data.
     *
     * @since 2.1.1
     */
    @Test
    public void listFill() {
        // Template note: Use {} to represent variables. If the template contains "{", "}" as special characters, use "\{", "\}" instead.
        // When filling a list, note that {.} in the template indicates a list.
        String templateFileName =
            TestFileUtil.getPath() + "demo" + File.separator + "fill" + File.separator + "list.xlsx";

        // Option 1: Load all data into memory at once and fill
        String fileName = TestFileUtil.getPath() + "listFill" + System.currentTimeMillis() + ".xlsx";
        // This will fill the first sheet, and the file stream will be closed automatically.
        EasyExcel.write(fileName).withTemplate(templateFileName).sheet().doFill(data());

        // Option 2: Fill in multiple passes using file caching (saves memory)
        fileName = TestFileUtil.getPath() + "listFill" + System.currentTimeMillis() + ".xlsx";
        ExcelWriter excelWriter = EasyExcel.write(fileName).withTemplate(templateFileName).build();
        WriteSheet writeSheet = EasyExcel.writerSheet().build();
        excelWriter.fill(data(), writeSheet);
        excelWriter.fill(data(), writeSheet);
        // Do not forget to close the stream
        excelWriter.finish();
    }

    /**
     * Example of complex data filling.
     *
     * @since 2.1.1
     */
    @Test
    public void complexFill() {
        // Template note: Use {} to represent variables. If the template contains "{", "}" as special characters, use "\{", "\}" instead.
        // {} represents a regular variable, {.} represents a list variable.
        String templateFileName =
            TestFileUtil.getPath() + "demo" + File.separator + "fill" + File.separator + "complex.xlsx";

        String fileName = TestFileUtil.getPath() + "complexFill" + System.currentTimeMillis() + ".xlsx";
        ExcelWriter excelWriter = EasyExcel.write(fileName).withTemplate(templateFileName).build();
        WriteSheet writeSheet = EasyExcel.writerSheet().build();
        // Note: The `forceNewRow` parameter ensures that a new row is created when writing a list, even if there are no empty rows below it.
        // By default, it is false, meaning it will use the next row if available, or create one if not.
        // Setting `forceNewRow=true` has the drawback of loading all data into memory, so use it cautiously.
        // If your template has a list and data below it, you must set `forceNewRow=true`, but this will consume more memory.
        // For large datasets where the list is not the last row, refer to the next example.
        FillConfig fillConfig = FillConfig.builder().forceNewRow(Boolean.TRUE).build();
        excelWriter.fill(data(), fillConfig, writeSheet);
        excelWriter.fill(data(), fillConfig, writeSheet);
        Map<String, Object> map = new HashMap<String, Object>();
        map.put("date", "2019-10-09 13:28:28");
        map.put("total", 1000);
        excelWriter.fill(map, writeSheet);
        excelWriter.finish();
    }

    /**
     * Example of complex data filling with large datasets.
     * <p>
     * The solution here is to ensure the list in the template is the last row, then append a table.
     * Note: Excel 2003 format is not supported and requires more memory.
     *
     * @since 2.1.1
     */
    @Test
    public void complexFillWithTable() {
        // Template note: Use {} to represent variables. If the template contains "{", "}" as special characters, use "\{", "\}" instead.
        // {} represents a regular variable, {.} represents a list variable.
        // Here, the template removes data after the list, such as summary rows.
        String templateFileName =
            TestFileUtil.getPath() + "demo" + File.separator + "fill" + File.separator + "complexFillWithTable.xlsx";

        String fileName = TestFileUtil.getPath() + "complexFillWithTable" + System.currentTimeMillis() + ".xlsx";
        ExcelWriter excelWriter = EasyExcel.write(fileName).withTemplate(templateFileName).build();
        WriteSheet writeSheet = EasyExcel.writerSheet().build();
        // Directly write data
        excelWriter.fill(data(), writeSheet);
        excelWriter.fill(data(), writeSheet);

        // Write data before the list
        Map<String, Object> map = new HashMap<String, Object>();
        map.put("date", "2019-10-09 13:28:28");
        excelWriter.fill(map, writeSheet);

        // Manually write summary data after the list
        // Here, we use a list for simplicity, but an object could also be used.
        List<List<String>> totalListList = new ArrayList<List<String>>();
        List<String> totalList = new ArrayList<String>();
        totalListList.add(totalList);
        totalList.add(null);
        totalList.add(null);
        totalList.add(null);
        // Fourth column
        totalList.add("Total: 1000");
        // Use `write` instead of `fill` here
        excelWriter.write(totalListList, writeSheet);
        excelWriter.finish();
        // Overall, this approach is complex, but no better solution is available.
        // Asynchronous writing to Excel does not support row deletion or movement, nor does it support comments.
        // Therefore, this workaround is necessary.
    }

    /**
     * Example of horizontal data filling.
     *
     * @since 2.1.1
     */
    @Test
    public void horizontalFill() {
        // Template note: Use {} to represent variables. If the template contains "{", "}" as special characters, use "\{", "\}" instead.
        // {} represents a regular variable, {.} represents a list variable.
        String templateFileName =
            TestFileUtil.getPath() + "demo" + File.separator + "fill" + File.separator + "horizontal.xlsx";

        String fileName = TestFileUtil.getPath() + "horizontalFill" + System.currentTimeMillis() + ".xlsx";
        ExcelWriter excelWriter = EasyExcel.write(fileName).withTemplate(templateFileName).build();
        WriteSheet writeSheet = EasyExcel.writerSheet().build();
        FillConfig fillConfig = FillConfig.builder().direction(WriteDirectionEnum.HORIZONTAL).build();
        excelWriter.fill(data(), fillConfig, writeSheet);
        excelWriter.fill(data(), fillConfig, writeSheet);

        Map<String, Object> map = new HashMap<String, Object>();
        map.put("date", "2019-10-09 13:28:28");
        excelWriter.fill(map, writeSheet);

        // Do not forget to close the stream
        excelWriter.finish();
    }

    /**
     * Example of composite data filling with multiple lists.
     *
     * @since 2.2.0-beta1
     */
    @Test
    public void compositeFill() {
        // Template note: Use {} to represent variables. If the template contains "{", "}" as special characters, use "\{", "\}" instead.
        // {} represents a regular variable, {.} represents a list variable, and {prefix.} distinguishes different lists.
        String templateFileName =
            TestFileUtil.getPath() + "demo" + File.separator + "fill" + File.separator + "composite.xlsx";

        String fileName = TestFileUtil.getPath() + "compositeFill" + System.currentTimeMillis() + ".xlsx";
        ExcelWriter excelWriter = EasyExcel.write(fileName).withTemplate(templateFileName).build();
        WriteSheet writeSheet = EasyExcel.writerSheet().build();
        FillConfig fillConfig = FillConfig.builder().direction(WriteDirectionEnum.HORIZONTAL).build();
        // If there are multiple lists, the template must use {prefix.}, where "data1" is the prefix.
        // Multiple lists must be wrapped in FillWrapper.
        excelWriter.fill(new FillWrapper("data1", data()), fillConfig, writeSheet);
        excelWriter.fill(new FillWrapper("data1", data()), fillConfig, writeSheet);
        excelWriter.fill(new FillWrapper("data2", data()), writeSheet);
        excelWriter.fill(new FillWrapper("data2", data()), writeSheet);
        excelWriter.fill(new FillWrapper("data3", data()), writeSheet);
        excelWriter.fill(new FillWrapper("data3", data()), writeSheet);

        Map<String, Object> map = new HashMap<String, Object>();
        //map.put("date", "2019-10-09 13:28:28");
        map.put("date", new Date());

        excelWriter.fill(map, writeSheet);

        // Do not forget to close the stream
        excelWriter.finish();
    }

    /**
     * Generates sample data for filling.
     *
     * @return A list of FillData objects.
     */
    private List<FillData> data() {
        List<FillData> list = new ArrayList<FillData>();
        for (int i = 0; i < 10; i++) {
            FillData fillData = new FillData();
            list.add(fillData);
            fillData.setName("Zhang San");
            fillData.setNumber(5.2);
        }
        return list;
    }
}
