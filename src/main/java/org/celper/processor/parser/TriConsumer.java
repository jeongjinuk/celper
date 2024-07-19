package org.celper.processor.parser;

public interface TriConsumer<T,U,R> {
    void accept(T t, U u, R r);
}
