package org.celper;

import org.celper.core.style.CellStyleConfigurer;
import org.celper.core.style.SheetStyleConfigurer;
import org.celper.register.RegisterManager;

import java.util.Arrays;
import java.util.Collections;
import java.util.List;
import java.util.function.Function;

public final class Generated__MetaData implements MetaData<DTO> {


    /** [0] : Type Class<?>
     *  [1] : Type Class<?>
     * private static final MetaData<[0]> INSTANCE = new [1]();
     */
    private static final MetaData<DTO> INSTANCE = new Generated__MetaData();

    private static final SheetStyleConfigurer sheetStyle = builder -> {};
    private static final List<CellStyleConfigurer> headerStyleList = Collections.unmodifiableList(Arrays.asList());
    private static final List<CellStyleConfigurer> dataStyleList = Collections.unmodifiableList(Arrays.asList());
    private static final List<String> columnNameList = Collections.unmodifiableList(Arrays.asList());
    private static final List<Object> defaultValueList = Collections.unmodifiableList(Arrays.asList());
    private static final List<Function<DTO, Object>> getterList = Collections.unmodifiableList(Arrays.asList(DTO::getAge, ...));

    /**
     * [0] : Type Class<?>
     * RegisterManager.put( [0] , INSTANCE);
     */
    static {
        RegisterManager.put(DTO.class, INSTANCE);
    }

    private Generated__MetaData() {}

    /**
     * 불변
     */
    @Override
    public SheetStyleConfigurer getSheetStyleConfig() {
        return this.sheetStyle;
    }

    @Override
    public List<CellStyleConfigurer> getHeaderStyleConfigs() {
        return this.headerStyleList;
    }

    @Override
    public List<CellStyleConfigurer> getDataStyleConfigs() {
        return this.dataStyleList;
    }

    @Override
    public List<String> getColumnNames() {
        return this.columnNameList;
    }

    @Override
    public List<Object> getDefaultValues() {
        return this.defaultValueList;
    }

    @Override
    public List<Function<DTO, Object>> getGetterFunctions() {
        return this.getterList;
    }
}
