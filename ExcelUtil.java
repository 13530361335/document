import com.alibaba.fastjson.JSONObject;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddressList;
import org.apache.poi.xssf.streaming.SXSSFCell;
import org.apache.poi.xssf.streaming.SXSSFRow;
import org.apache.poi.xssf.streaming.SXSSFSheet;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFDataValidationConstraint;
import org.apache.poi.xssf.usermodel.XSSFDataValidationHelper;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.springframework.stereotype.Component;

import javax.servlet.http.HttpServletResponse;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.text.SimpleDateFormat;
import java.util.LinkedHashMap;
import java.util.LinkedList;
import java.util.List;
import java.util.Map.Entry;

@Component
public class ExcelUtil {

    private static Logger logger = LoggerFactory.getLogger(ExcelUtil.class);

    /**
     * 导出excel
     *
     * @param head     列头，key为数据字段名称，value为excel字段名称
     * @param datas    数据集
     * @param response
     * @param fileName 导出文件名称
     * @throws IOException
     */
    public static void exportXlsx(LinkedHashMap<String, String> head, List<JSONObject> datas, HttpServletResponse response,
                                  String fileName) throws IOException {
        // 声明一个工作薄,生成一个表格
        SXSSFWorkbook workbook = new SXSSFWorkbook(1000);
        workbook.setCompressTempFiles(true);
        SXSSFSheet sheet = workbook.createSheet();

        // 设置列头和字段对应关系
        String[] Fields = new String[head.size()];// 字段名称
        String[] excelFields = new String[head.size()];// excel字段名称
        int colIndex = 0;
        for (Entry<String, String> entry : head.entrySet()) {
            Fields[colIndex] = entry.getKey();
            excelFields[colIndex] = entry.getValue();
            colIndex++;
        }

        // 加载列头
        SXSSFRow headRow = sheet.createRow(0);
        sheet.trackAllColumnsForAutoSizing();// 列宽自适应
        for (int i = 0; i < excelFields.length; i++) {
            headRow.createCell(i).setCellValue(excelFields[i]);
            sheet.autoSizeColumn(i);// 列宽自适应
        }

        // 加载数据行
        int rowIndex = 1;
        for (JSONObject data : datas) {
            SXSSFRow dataRow = sheet.createRow(rowIndex);
            for (int i = 0; i < Fields.length; i++) {
                SXSSFCell cell = dataRow.createCell(i);
                Object value = data.get(Fields[i]);
                String cellValue;
                cellValue = value == null ? "" : value.toString();
                cell.setCellValue(cellValue);
            }
            rowIndex++;
        }

        // 文件传输
        response.setContentType("application/octet-stream");
        response.setHeader("Content-Disposition",
                "attachment;filename=" + new String(fileName.getBytes("UTF-8"), "ISO8859-1") + ".xlsx");
        OutputStream out = response.getOutputStream();
        workbook.write(out);
        workbook.close();
        out.close();
    }

    /**
     * excel转化为json数据
     *
     * @param fileds    字段数组
     * @param in
     * @param startLine 数据起始行，第一行为0
     * @return
     * @throws IOException [参数说明]
     */
    public static List<JSONObject> parseXlsx(String[] fileds, InputStream in, int startLine) throws IOException {
        // 从流中获取一个工作薄,获取第一个表格
        Workbook work = new XSSFWorkbook(in);
        Sheet sheet = work.getSheetAt(0);

        // 转换之后的数据
        List<JSONObject> datas = new LinkedList();
        for (int rowIndex = startLine; rowIndex <= sheet.getLastRowNum(); rowIndex++) {
            Row row = sheet.getRow(rowIndex);
            JSONObject data = new JSONObject();
            for (int colIndex = row.getFirstCellNum(); colIndex < fileds.length; colIndex++) {
                Cell cell = row.getCell(colIndex);
                String cellValue = getCellValue(cell);
                data.put(fileds[colIndex], cellValue);
            }
            datas.add(data);
        }

        try {
            work.close();
        } catch (IOException e) {
            logger.error(e.getMessage(), e);
        }
        return datas;
    }

    private static String getCellValue(Cell cell) {
        String cellValue = "";
        if (cell != null) {
            switch (cell.getCellType()) {
                case Cell.CELL_TYPE_NUMERIC: // 数字
                    if (DateUtil.isCellDateFormatted(cell)) {
                        double numValue = cell.getNumericCellValue();
                        SimpleDateFormat sdf = new SimpleDateFormat("yyyy-MM-dd");
                        cellValue = sdf.format(DateUtil.getJavaDate(numValue));
                    } else {
                        DataFormatter dataFormatter = new DataFormatter();
                        cellValue = dataFormatter.formatCellValue(cell);
                    }
                    break;
                case Cell.CELL_TYPE_STRING: // 字符串
                    cellValue = cell.getStringCellValue();
                    break;
                case Cell.CELL_TYPE_BOOLEAN: // Boolean
                    cellValue = cell.getBooleanCellValue() + "";
                    break;
                case Cell.CELL_TYPE_FORMULA: // 公式
                    cellValue = cell.getCellFormula() + "";
                    break;
                case Cell.CELL_TYPE_BLANK: // 空值
                    cellValue = "";
                    break;
                case Cell.CELL_TYPE_ERROR: // 故障
                    cellValue = "非法字符";
                    break;
                default:
                    cellValue = "未知类型";
                    break;
            }
        }
        return cellValue;
    }

    // 设置下拉框
    public static void setDropDownBox(XSSFSheet sheet, String[] datas, int colIndex) {
        XSSFDataValidationHelper dvHelper = new XSSFDataValidationHelper(sheet);
        XSSFDataValidationConstraint constraint = (XSSFDataValidationConstraint) dvHelper
                .createExplicitListConstraint(datas);
        CellRangeAddressList addressList = new CellRangeAddressList(1, 100, colIndex, colIndex);
        DataValidation validation = dvHelper.createValidation(constraint, addressList);
        sheet.addValidationData(validation);
    }
}
