package org.celper;

import org.celper.core.style.*;
import org.celper.register.RegisterManager;

import java.util.*;
import java.util.function.*;

public final class Generated_MetaData implements MetaData<DTO> {
    private static final MetaData<DTO> INSTANCE = new Generated_MetaData();

    // new $T
    private static final SheetStyleConfigurer sheetStyle = builder -> {};

    // new $T
    private static final List<CellStyleConfigurer> headerStyleList = Collections.unmodifiableList(Arrays.asList());
    // new $T
    private static final List<CellStyleConfigurer> dataStyleList = Collections.unmodifiableList(Arrays.asList());
    private static final List<String> columnNameList = Collections.unmodifiableList(Arrays.asList(" ","1", "3", "4"));
    private static final List<Object> defaultValueList = Collections.unmodifiableList(Arrays.asList(new Object()));
    private static final List<Function<DTO, Object>> getterList = Collections.unmodifiableList(Arrays.asList(DTO::getAge, DTO::getGrade));

    static {
        RegisterManager.put(DTO.class, INSTANCE);
    }

    private Generated_MetaData() {}
    @Override
    public SheetStyleConfigurer getSheetStyleConfig() {
        return sheetStyle;
    }

    @Override
    public List<CellStyleConfigurer> getHeaderStyleConfigs() {
        return headerStyleList;
    }

    @Override
    public List<CellStyleConfigurer> getDataStyleConfigs() {
        return dataStyleList;
    }

    @Override
    public List<String> getColumnNames() {
        return columnNameList;
    }

    @Override
    public List<Object> getDefaultValues() {
        return defaultValueList;
    }

    @Override
    public List<Function<DTO, Object>> getGetterFunctions() {
        return getterList;
    }
}
