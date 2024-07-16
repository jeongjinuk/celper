package org.celper;

import org.celper.core.style.*;
import org.celper.register.RegisterManager;

import java.util.*;
import java.util.function.*;

public final class Generated_MetaData implements MetaData<DTO> {
    private static final MetaData<DTO> INSTANCE = new Generated_MetaData();
    private final SheetStyleConfigurer SheetStyle = builder -> {};
    private final List<CellStyleConfigurer> headerStyleConfigList = Collections.unmodifiableList(Arrays.asList());
    // new $
    private final List<CellStyleConfigurer> dataStyleConfigList = Collections.unmodifiableList(Arrays.asList());
    private final List<String> columnNameList = Collections.unmodifiableList(Arrays.asList(" ","1", "3", "4"));
    private final List<String> defaultValueList = Collections.unmodifiableList(Arrays.asList(" ", null));
    private final List<String> cellFormatList = Collections.unmodifiableList(Arrays.asList(" ", null));
    private final List<Function<DTO, Object>> getterFunctionList = Collections.unmodifiableList(Arrays.asList(DTO::getAge, DTO::getGrade));

    static {
        RegisterManager.put(DTO.class, INSTANCE);
    }

    @Override
    public SheetStyleConfigurer getSheetStyle() {
        return SheetStyle;
    }

    @Override
    public List<CellStyleConfigurer> getHeaderStyleConfigList() {
        return headerStyleConfigList;
    }

    @Override
    public List<CellStyleConfigurer> getDataStyleConfigList() {
        return dataStyleConfigList;
    }

    @Override
    public List<String> getColumnNameList() {
        return columnNameList;
    }

    @Override
    public List<String> getDefaultValueList() {
        return defaultValueList;
    }

    @Override
    public List<String> getCellFormatList() {
        return cellFormatList;
    }

    @Override
    public List<Function<DTO, Object>> getGetterFunctionList() {
        return getterFunctionList;
    }
}
