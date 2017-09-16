package tv.lycam.excel.utils.cglib;

import com.alibaba.fastjson.JSON;
import com.alibaba.fastjson.JSONObject;
import org.apache.poi.hssf.usermodel.HSSFCellStyle;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import tv.lycam.excel.model.CglibBean;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.lang.reflect.Field;
import java.util.*;

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
    public static List<Object> importExcel(String file, int sheetIndex) throws IOException {
        List<Object> result = new ArrayList<Object>();

        if (file == null || file.isEmpty()) {
            return result;
        }
        FileInputStream in = null;
        try {
            in = new FileInputStream(file);
            Workbook wb;
            if (file.endsWith(".xlsx")) {
                wb = new XSSFWorkbook(in);
            } else {
                wb = new HSSFWorkbook(in);
            }
            Sheet sheet = wb.getSheetAt(sheetIndex);
            Set<String> fieldNames = CglibUtils.getPropertyMap().keySet();
            CglibBean bean;
            for (Row row : sheet) {
                if (row.getRowNum() < 1) {
                    if (fieldNames.size() == 0) {
                        short firstCellNum = row.getFirstCellNum();
                        short lastCellNum = row.getLastCellNum();
                        fieldNames = new LinkedHashSet<>();
                        for (int i = 0; i < lastCellNum; i++) {
                            Cell cell = row.getCell(i);
                            String fieldName = cell.getStringCellValue();
                            fieldNames.add(fieldName);
                        }
                        CglibUtils.setPropertyMap(fieldNames);
                    }
                    continue;
                }
                int rowNum = 0;
                bean = CglibUtils.generateBean();
                Object o = bean.getObject();
                for (String fieldName : fieldNames) {
                    Cell cell = row.getCell(rowNum);
                    if (cell == null) {
                        continue;
                    }
                    Field field = PropertyUtils.getFieldByName(o.getClass(), fieldName);
                    if (field.getType() == Boolean.class) {
                        boolean value = cell.getBooleanCellValue();
                        bean.setValue(fieldName, value);
                    } else if (field.getType() == Date.class) {
                        Date value = cell.getDateCellValue();
                        bean.setValue(fieldName, value);
                    } else if (field.getType() == String.class) {
                        String value = cell.getStringCellValue();
                        bean.setValue(fieldName, value);
                    } else if (field.getType() == Double.class) {
                        double value = cell.getNumericCellValue();
                        bean.setValue(fieldName, value);
                    } else if (field.getType() == RichTextString.class) {
                        RichTextString value = cell.getRichStringCellValue();
                        bean.setValue(fieldName, value);
                    } else {
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

    public static void exportAsExcel(List<String> jsons, String path, String collection) throws IOException {
        if (jsons == null || jsons.size() == 0) {
            return;
        }
        HSSFWorkbook wb = new HSSFWorkbook();

        HSSFSheet sheet = wb.createSheet(collection);
        HSSFCellStyle style = wb.createCellStyle();
        style.setAlignment(HorizontalAlignment.CENTER);
        int rowCount = 0;

        Set<String> fieldNames = CglibUtils.getPropertyMap().keySet();
        if (fieldNames.size() == 0) {
            JSONObject jobj = JSONObject.parseObject(jsons.get(0));
            fieldNames = jobj.keySet();
            CglibUtils.setPropertyMap(fieldNames);
        }
        HSSFRow firstRow = sheet.createRow(0);
        for (String key : fieldNames) {
            Cell cell = firstRow.createCell(rowCount);
            cell.setCellStyle(style);
            cell.setCellValue(key);
            rowCount++;
            System.out.println(key);
        }
        for (int i = 0; i < jsons.size(); i++) {
            HSSFRow row = sheet.createRow(i + 1);
            String json = jsons.get(i);
            CglibBean bean = CglibUtils.generateBean();
            Object obj = JSON.parseObject(json, bean.getObject().getClass());
            bean.setObject(obj);
            rowCount = 0;
            for (String key : fieldNames) {
                Cell cell = row.createCell(rowCount);
                cell.setCellStyle(style);
                Object value = bean.getValue(key);
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
//                System.out.println(String.format("%s:\t%s", key, String.valueOf(o)));
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

    public static List<Object> importExcel(String file) throws IOException {
        return importExcel(file, 0);
    }
}
