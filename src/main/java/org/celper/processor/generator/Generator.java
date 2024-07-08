package org.celper.processor.generator;

public interface Generator<T, R> {
    R generate(T t);
}
