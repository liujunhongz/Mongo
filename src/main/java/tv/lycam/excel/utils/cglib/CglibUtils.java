package tv.lycam.excel.utils.cglib;

import tv.lycam.excel.model.CglibBean;

import java.io.FileInputStream;
import java.io.IOException;
import java.util.HashMap;
import java.util.LinkedHashMap;
import java.util.Properties;
import java.util.Set;

/**
 * Cglib工具类
 *
 * @author cuiran
 * @version 1.0
 */
public class CglibUtils {

    private static HashMap<String, Class> propertyMap;
    public static String NotNull;

    @SuppressWarnings("unchecked")
    public static HashMap<String, Class> getPropertyMap() {
        try {
            if (propertyMap != null) {
                return propertyMap;
            }
            propertyMap = new LinkedHashMap<>();
            Properties prop = new Properties();
            prop.load(new FileInputStream("config.properties"));
            NotNull = prop.getProperty("NotNull", "");
            Set<String> propertyNames = prop.stringPropertyNames();
            // 设置类成员属性
            for (String name : propertyNames) {
                if ("NotNull".equalsIgnoreCase(name)) {
                    continue;
                }
                propertyMap.put(name, Class.forName(prop.getProperty(name)));
            }
            return propertyMap;

        } catch (IOException | ClassNotFoundException e) {
            throw new RuntimeException(e);
        }

    }

    public static CglibBean generateBean() {
        // 生成动态 Bean
        CglibBean bean = new CglibBean(getPropertyMap());
        return bean;
    }

    public static void setPropertyMap(Set<String> fieldName) {
        for (String name : fieldName) {
            propertyMap.put(name, Object.class);
        }
    }
}