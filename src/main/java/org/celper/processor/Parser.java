package org.celper.processor;

public interface Parser<T,R> {
    R parse(T t);
}
