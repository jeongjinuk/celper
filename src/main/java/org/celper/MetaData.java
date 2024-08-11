package org.celper;

import org.celper.core2.style.CellStyleConfigurer;
import org.celper.core2.style.SheetLayoutConfigurer;

import java.util.List;
import java.util.function.Function;

/**
 * CodeBlock.Builder cb = $T.unmodifiableList($T.asList(
 * cb.add($S) or cb.add(new $T()
 *
 */
public interface MetaData<T> {
    SheetLayoutConfigurer getSheetLayout();
    CellStyleConfigurer getSheetStyle();
    List<CellStyleConfigurer> getHeaderStyleConfigList();
    List<CellStyleConfigurer> getDataStyleConfigList();
    List<String> getColumnNameList();
    List<String> getDefaultValueList();
    List<String> getCellFormatList();
    List<Function<T, Object>> getGetterFunctionList();
    String[] getCsvConfig();
}
