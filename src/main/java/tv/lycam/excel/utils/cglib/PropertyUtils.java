package tv.lycam.excel.utils.cglib;

import java.lang.reflect.Field;
import java.lang.reflect.Method;

public class PropertyUtils {
    private static String CGLIB = "$cglib_prop_";

    /**
     * 获取属性名数组
     */
    public static String[] getFieldNames(Class clazz) {
        Field[] fields = clazz.getDeclaredFields();
        String[] fieldNames = new String[fields.length];
        for (int i = 0; i < fields.length; i++) {
            fieldNames[i] = fields[i].getName();
        }
        return fieldNames;
    }

    public static Field getFieldByName(Class clazz, String fieldName) {
        try {
            Field field = clazz.getDeclaredField(CGLIB + fieldName);
            return field;
        } catch (NoSuchFieldException e) {
            throw new RuntimeException(e);
        }
    }

    /**
     * 设置属性值
     */
    public static void setCglibFieldValue(Object o, String fieldName, Object fieldValue) {
        try {
            Field field = o.getClass().getDeclaredField(fieldName);

            String firstLetter = fieldName.replace(CGLIB, "").substring(0, 1).toUpperCase();
            String setter = "set" + firstLetter + fieldName.substring(1);
            Method method = o.getClass().getMethod(setter, field.getType());
            method.invoke(o, fieldValue);
        } catch (Exception e) {
            e.printStackTrace();
        }
    }

    /**
     * 根据属性名获取属性值
     */
    public static Object getFieldValueByName(String fieldName, Object o) {
        try {
            String firstLetter = fieldName.replace(CGLIB, "").substring(0, 1).toUpperCase();
            String getter = "get" + firstLetter + fieldName.substring(1);
            Method method = o.getClass().getMethod(getter);
            return method.invoke(o);
        } catch (Exception e) {
            e.printStackTrace();
            return null;
        }
    }
}
