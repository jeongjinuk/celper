package org.celper.register;

import io.github.classgraph.ClassGraph;
import io.github.classgraph.ScanResult;
import org.celper.MetaData;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.util.concurrent.ConcurrentHashMap;

public final class RegisterManager {

    private static final Logger log = LoggerFactory.getLogger(RegisterManager.class);
    private static final ConcurrentHashMap<Class<?>, MetaData<?>> REGISTER = new ConcurrentHashMap<>();

    static {
        try (ScanResult result = new ClassGraph().enableClassInfo().scan()) {
            result.getClassesImplementing(MetaData.class.getName())
                    .loadClasses()
                    .forEach(RegisterManager::initializeClass);
        }
    }
    // TODO Lazy initialize 형태로?
    public static void initializeClass(Class<?> clazz) {
        try{
            Class.forName(clazz.getName(), true, clazz.getClassLoader());
            log.info("initialized Class : [" + clazz.getName() + "]");
        } catch (ClassNotFoundException e) {
            log.error("failed initialize Class : [" + clazz.getName() + "]", e);
        }
    }
    // TODO Optional.ofNullable() 고려
    @SuppressWarnings("unchecked")
    public static <T> MetaData<T> getMetaData(Class<T> clazz) {
        return (MetaData<T>) REGISTER.get(clazz);
    }

    public static void register(Class<?> clazz, MetaData<?> generatedMetaData) {
        REGISTER.putIfAbsent(clazz, generatedMetaData);
    }

}