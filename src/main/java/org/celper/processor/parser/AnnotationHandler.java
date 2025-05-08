package org.celper.processor.parser;

import org.celper.processor.util.ElementUtil;

public interface AnnotationHandler<T,S> {

    static <U,R> void applyHandlersIfPresent(AnnotationHandler<U, R>[] handlers,
                                              ElementUtil elementUtil,
                                              U u,
                                              R r){
        for (AnnotationHandler<U,R> handler : handlers) {
            handler.ifPresent(u, r,elementUtil);
        }
    }
    void ifPresent(T t, S s,ElementUtil elementUtil);
}
