package org.celper.processor;

public interface Validator<T>{
    boolean valid(T t);
}
