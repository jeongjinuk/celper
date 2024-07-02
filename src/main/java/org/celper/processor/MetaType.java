package org.celper.processor;

import org.celper.core.style.CellStyleConfigurer;
import org.celper.core.style.SheetStyleConfigurer;

import javax.lang.model.element.Element;

public enum MetaType {
    FIELD(Element.class),
    SHEET_STYLE(SheetStyleConfigurer.class),
    HEADER_STYLE(CellStyleConfigurer.class),
    CELL_STYLE(CellStyleConfigurer.class),
    FORMAT(String.class),
    HEADER_NAME(String.class),
    DEFAULT_VALUE(String.class);
    private Class<?> clazz;

    MetaType(Class<?> clazz) {
        this.clazz = clazz;
    }

    public Class<?> getClazz() {
        return clazz;
    }
}
