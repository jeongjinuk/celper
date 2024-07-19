package org.celper.register;

import io.github.classgraph.ClassGraph;
import io.github.classgraph.ClassInfo;
import io.github.classgraph.ClassInfoList;
import org.celper.MetaData;

import java.util.concurrent.ConcurrentHashMap;

public class RegisterManager {
    private static final ConcurrentHashMap<Class<?>, MetaData<?>> REGISTER = new ConcurrentHashMap<>();

    static {
        init();
    }
    public static <T>  MetaData<T> getMetaData(Class<T> clazz) {
        return (MetaData<T>) REGISTER.get(clazz);
    }

    static void init() {
        ClassInfoList classesImplementing = new ClassGraph().enableAllInfo()
                .scan()
                .getClassesImplementing(MetaData.class.getName());
        for (ClassInfo classInfo : classesImplementing) {
            Class<?> aClass = classInfo.loadClass();
            try {
                Class.forName(aClass.getName(), true, aClass.getClassLoader());
            } catch (ClassNotFoundException ignore) {
                // nothing
            }
        }
    }

    // TODO generated Model 이름 변경
    public static void put(Class<?> clazz, MetaData<?> generatedMetaData){
        REGISTER.put(clazz, generatedMetaData);
    }

}
