package tv.lycam.excel.model;

import java.lang.reflect.InvocationTargetException;
import java.lang.reflect.Method;

public class CglibObject {

    @Override
    public String toString() {

        StringBuilder builder = new StringBuilder();
        Class clazz = getClass();
        Method[] methods = clazz.getDeclaredMethods();
        for (Method method : methods) {
            if (method.getName().startsWith("get")) {
                try {
                    Object value = method.invoke(this);
                    builder.append("\t").append(value);
                } catch (IllegalAccessException | InvocationTargetException e) {
                    e.printStackTrace();
                }
            }
        }
        return builder.toString();
    }
}
