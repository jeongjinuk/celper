package org.celper.processor.parser;

public interface AnnotationHandler<T,S> {

    static  <U,R> void applyHandlersIfPresent(AnnotationHandler<U, R>[] handlers, U u, R r){
        for (AnnotationHandler<U,R> handler : handlers) {
            handler.ifPresent(u, r);
        }
    }
    void ifPresent(T t, S s);
}
