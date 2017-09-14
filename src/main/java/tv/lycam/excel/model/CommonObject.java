package tv.lycam.excel.model;

import java.lang.reflect.InvocationTargetException;
import java.lang.reflect.Method;

public class CommonObject {

    public String displayName;
    public String phone;

    public String getDisplayName() {
        return displayName;
    }

    public void setDisplayName(String displayName) {
        this.displayName = displayName;
    }

    public String getPhone() {
        return phone;
    }

    public void setPhone(String phone) {
        this.phone = phone;
    }

    @Override
    public String toString() {

        StringBuilder builder = new StringBuilder();
        Class clazz = getClass();
        Method[] methods = clazz.getDeclaredMethods();
        for (Method method : methods) {
            if (method.getName().startsWith("get")) {
                try {
                    java.lang.Object value = method.invoke(this);
                    builder.append("\t").append(value);
                } catch (IllegalAccessException | InvocationTargetException e) {
                    e.printStackTrace();
                }
            }
        }
        return builder.toString();
    }
}
