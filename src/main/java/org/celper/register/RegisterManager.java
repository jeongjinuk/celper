package org.celper.register;

import org.celper.MetaData;

import java.util.concurrent.ConcurrentHashMap;

public class RegisterManager {
    private static final ConcurrentHashMap<Class<?>, MetaData<?>> REGISTER = new ConcurrentHashMap<>();

    static {
        init();
    }
    public static void init(){
        // ClassGraph -> classInfo -> class.forName() -> load
    }

    // TODO generated Model 이름 변경
    public static void put(Class<?> clazz, MetaData<?> generatedMetaData){

    }


}
