package org.celper;

import org.celper.core.style.CellStyleConfigurer;
import org.celper.core.style.SheetStyleConfigurer;

import java.util.List;
import java.util.function.Function;

public interface Model<T> {
    SheetStyleConfigurer getSheetStyle();
    List<CellStyleConfigurer> getHeaderStyle();
    List<CellStyleConfigurer> getDataStyle();
    List<String> getColumnNames();
    List<Object> getDefaultValues();
    List<Function<T, Object>> getConsumers();
}
