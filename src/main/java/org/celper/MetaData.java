package org.celper;

import org.celper.core.style.CellStyleConfigurer;
import org.celper.core.style.SheetStyleConfigurer;

import java.util.List;
import java.util.function.Function;

public interface MetaData<T> {
    SheetStyleConfigurer getSheetStyleConfig();
    List<CellStyleConfigurer> getHeaderStyleConfigs();
    List<CellStyleConfigurer> getDataStyleConfigs();
    List<String> getColumnNames();
    List<Object> getDefaultValues();
    List<Function<T, Object>> getGetterFunctions();
}
