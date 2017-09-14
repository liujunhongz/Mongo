package tv.lycam.excel.utils.common;

import com.alibaba.fastjson.JSON;
import org.apache.poi.hssf.usermodel.HSSFCellStyle;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import tv.lycam.excel.model.CommonObject;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.lang.reflect.Field;
import java.util.ArrayList;
import java.util.Calendar;
import java.util.Date;
import java.util.List;

public class ExcelUtils {
    /**
     * 07之前版本
     *
     * @throws IOException
     */
    public static void excelBefore07() throws IOException {
        Workbook wb = new HSSFWorkbook();
        FileOutputStream fos = new FileOutputStream("workbook.xls");
        wb.write(fos);
        fos.close();
    }

    /**
     * 07之后版本
     *
     * @throws IOException
     */
    public static void excelAfter07() throws IOException {
        Workbook wb = new XSSFWorkbook();
        FileOutputStream fos = new FileOutputStream("workbook.xlsx");
        wb.write(fos);
        fos.close();
    }

    /**
     * 导入excel文件，使用绝对路径
     *
     * @param file
     * @param sheetIndex
     * @return
     * @throws IOException
     */
    public static List<CommonObject> importExcel(String file, int sheetIndex) throws IOException {
        FileInputStream in = null;
        List<CommonObject> result = null;
        try {
            in = new FileInputStream(file);
            result = new ArrayList<CommonObject>();
            Workbook wb = new HSSFWorkbook(in);
            Sheet sheet = wb.getSheetAt(sheetIndex);
            String[] fieldNames = PropertyUtils.getFieldNames(CommonObject.class);
            for (Row row : sheet) {
                if (row.getRowNum() < 1) {
                    continue;
                }
                int rowNum = 0;
                CommonObject o = new CommonObject();
                for (String fieldName : fieldNames) {
                    Cell cell = row.getCell(rowNum);
                    Field field = CommonObject.class.getDeclaredField(fieldName);
                    if (field.getType() == Boolean.class) {
                        boolean value = cell.getBooleanCellValue();
                        PropertyUtils.setFieldValue(o, fieldName, value);
                    } else if (field.getType() == Date.class) {
                        Date value = cell.getDateCellValue();
                        PropertyUtils.setFieldValue(o, fieldName, value);
                    } else if (field.getType() == String.class) {
                        String value = cell.getStringCellValue();
                        PropertyUtils.setFieldValue(o, fieldName, value);
                    } else if (field.getType() == Double.class) {
                        double value = cell.getNumericCellValue();
                        PropertyUtils.setFieldValue(o, fieldName, value);
                    } else if (field.getType() == RichTextString.class) {
                        RichTextString value = cell.getRichStringCellValue();
                        PropertyUtils.setFieldValue(o, fieldName, value);
                    }
                    rowNum++;
                }
                result.add(o);
            }
        } catch (Exception e) {
            e.printStackTrace();
        } finally {
            in.close();
        }
        return result;
    }

    public static void exportAsExcel(List<String> jsons, String path, String collection) {
        if (jsons == null || jsons.size() == 0) {
            return;
        }
        HSSFWorkbook wb = new HSSFWorkbook();

        HSSFSheet sheet = wb.createSheet(collection);
        HSSFCellStyle style = wb.createCellStyle();
        style.setAlignment(HorizontalAlignment.CENTER);
        int rowCount = 0;
        String[] fields = PropertyUtils.getFieldNames(CommonObject.class);

        HSSFRow firstRow = sheet.createRow(0);
        for (String key : fields) {
            Cell cell = firstRow.createCell(rowCount);
            cell.setCellStyle(style);
            cell.setCellValue(key);
            rowCount++;
            System.out.println(key);
        }
        for (int i = 0; i < jsons.size(); i++) {
            HSSFRow row = sheet.createRow(i + 1);
            String json = jsons.get(i);
            CommonObject obj = JSON.parseObject(json, CommonObject.class);
            rowCount = 0;
            for (String key : fields) {
                Cell cell = row.createCell(rowCount);
                cell.setCellStyle(style);
                Object value = PropertyUtils.getFieldValueByName(key, obj);
                if (value instanceof Boolean) {
                    cell.setCellValue((boolean) value);
                } else if (value instanceof Date) {
                    cell.setCellValue((Date) value);
                } else if (value instanceof Calendar) {
                    cell.setCellValue((Calendar) value);
                } else if (value instanceof String) {
                    cell.setCellValue((String) value);
                } else if (value instanceof Double || value instanceof Integer) {
                    cell.setCellValue((double) value);
                } else if (value instanceof RichTextString) {
                    cell.setCellValue((RichTextString) value);
                }
//                System.out.println(String.format("%s:\t%s", key, String.valueOf(obj)));
                rowCount++;
            }
        }
        try {
            FileOutputStream fout = new FileOutputStream(path);
            wb.write(fout);
            fout.flush();
            fout.close();
        } catch (Exception e) {
            e.printStackTrace();
        }
    }

    public static List<CommonObject> importExcel(String file) throws IOException {
        return importExcel(file, 0);
    }
}
