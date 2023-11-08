package org.celper.core;

import org.apache.poi.ss.usermodel.*;
import org.celper.core.model.ClassModel;
import org.celper.core.model.ColumnFrame;
import org.celper.exception.DataListEmptyException;
import org.celper.exception.NoSuchFieldException;
import org.celper.util.ModelMapperFactory;
import org.modelmapper.ModelMapper;

import java.util.*;
import java.util.function.Consumer;
import java.util.function.IntConsumer;
import java.util.function.Predicate;
import java.util.stream.Collectors;
import java.util.stream.IntStream;
import java.util.stream.Stream;
import java.util.stream.StreamSupport;

/**
 * The type Excel sheet.
 */
public class ExcelSheet {
    private final Workbook _wb;
    private final Sheet sheet;

    /**
     * Instantiates a new Excel sheet.
     *
     * @param workbook the workbook
     * @param sheet    the sheet
     */
    public ExcelSheet(Workbook workbook, Sheet sheet) {
        this._wb = workbook;
        this.sheet = sheet;
    }

    /**
     * Gets name.
     *
     * @return the name
     */
    public String getName() {
        return this.sheet.getSheetName();
    }

    /**
     * Gets sheet.
     *
     * @return the sheet
     */
    public Sheet getSheet() {
        return this.sheet;
    }

    /**
     * Model to sheet.
     * List가 비어있거나, null 배열일 경우 DataListEmptyException를 반환합니다.
     * @param <T>   the type parameter
     * @param model the model
     * @throws DataListEmptyException
     */
    public <T> void modelToSheet(List<T> model){
        modelToSheet(header -> false, model);
    }

    /**
     * Model to sheet.
     * List가 비어있거나, null 배열일 경우 DataListEmptyException를 반환합니다.
     * @param <T>            the type parameter
     * @param excludedHeader the excluded header
     * @param model          the model
     * @throws DataListEmptyException
     */
    public <T> void modelToSheet(Predicate<String> excludedHeader, List<T> model) {
        if (Objects.isNull(model) || model.isEmpty()) {
            throw new DataListEmptyException("data list is empty exception");
        }
        List<ColumnFrame> columnFrames = createColumnFrames(excludedHeader, ClassModelRegistrator.getOrDefault(model.get(0).getClass()));

        int headerRowIndex = 0;
        int dataRow = headerRowIndex + 1;
        int modelSize = model.size();

        headerWrite(columnFrames, headerRowIndex);
        IntStream.rangeClosed(dataRow, modelSize)
                .forEach(rowIndex -> dataWrite(columnFrames, rowIndex, model.get(rowIndex - 1)));
    }

    /**
     * Multi model to sheet.
     * null 배열일 경우 DataListEmptyException를 반환합니다.
     * @param modelLists the model lists
     * @throws DataListEmptyException
     */
    public void multiModelToSheet(List<?>... modelLists) {
        multiModelToSheet(s -> false, modelLists);
    }

    /**
     * Multi model to sheet.
     * null 배열일 경우 DataListEmptyException를 반환합니다.
     * @param excludedHeader the excluded header
     * @param modelLists     the model lists
     * @throws DataListEmptyException
     */
    public void multiModelToSheet(Predicate<String> excludedHeader, List<?>... modelLists) {
        for (List<?> modelList : modelLists) {
            if (Objects.isNull(modelList)) {
                throw new DataListEmptyException("data list is empty exception");
            }
        }

        List<Object[]> multiModel = convertMultiModel(modelLists);
        List<ColumnFrame> multiColumnFrames = createMultiColumnFrames(excludedHeader, multiModel);

        int headerRowIndex = 0;
        int dataRow = headerRowIndex + 1;
        int multiModelSize = multiModel.size();

        headerWrite(multiColumnFrames, headerRowIndex);
        IntStream.rangeClosed(dataRow, multiModelSize)
                .forEach(rowIndex -> multiModelDataWrite(multiColumnFrames, rowIndex, multiModel.get(rowIndex - 1)));
    }

    /**
     * Sheet to model list.
     * 만약 매칭되는 필드가 존재하지 않을 경우 NoSuchFieldException을 반환합니다.
     *
     * @param <T>   the type parameter
     * @param clazz the clazz
     * @return the list
     * @throws NoSuchFieldException
     *
     */
    public <T> List<T> sheetToModel(Class<T> clazz) {
        ModelMapper modelMapper = ModelMapperFactory.defaultModelMapper(); // Model Mapper 가져오기

        List<ColumnFrame> columnFrames = createColumnFrames(header -> false, ClassModelRegistrator.getOrDefault(clazz));
        List<ColumnFrame> importList = convertImportColumnFrames(columnFrames);

        int startRow = importList.stream()// 병합을 고려해서 startRow 추출 병합 고려할려면 isMerged 만들어야함
                .max(ColumnFrame :: compareRowPosition)
                .orElseThrow(() -> new NoSuchFieldException(String.format("'%s' 시트에 매칭된 필드가 없습니다.", this.sheet.getSheetName())))
                .getHeaderRowPosition() + 1;

        return IntStream.rangeClosed(startRow, this.sheet.getLastRowNum())
                .filter(rowIdx -> Objects.nonNull(this.sheet.getRow(rowIdx)))
                .mapToObj(rowIdx -> createModelMap(importList, rowIdx))
                .map(entry -> modelMapper.map(entry, clazz))
                .collect(Collectors.toList());
    }

    // render로 바꾸는게 좋아보임
    private void headerWrite(List<ColumnFrame> columnFrames, int rowIndex) {
        Row headerRow = CellUtils.createRow(this.sheet, rowIndex, columnFrames.size());
        IntConsumer setHeader = colIdx -> CellUtils.setValue(headerRow.getCell(colIdx), columnFrames.get(colIdx).getClassModel().getColumn().value());
        IntConsumer setStyle = colIdx -> headerRow.getCell(colIdx).setCellStyle(columnFrames.get(colIdx).getHeaderAreaCellStyle());
        write(columnFrames, setHeader, setStyle);
    }

    private void multiModelDataWrite(List<ColumnFrame> columnFrames, int rowIndex, Object[] model) {
        Stream.of(model).forEach(o -> dataWrite(columnFrames, rowIndex, o));
    }

    private void dataWrite(List<ColumnFrame> columnFrames, int rowIndex, Object o) {
        Row row = CellUtils.createRow(this.sheet, rowIndex, columnFrames.size());
        IntConsumer setValue = colIdx -> CellUtils.setValue(columnFrames.get(colIdx), row.getCell(colIdx), o);
        IntConsumer setStyle = colIdx -> row.getCell(colIdx).setCellStyle(columnFrames.get(colIdx).getDataAreaCellStyle());
        write(columnFrames, setValue, setStyle);
    }

    private void write(List<ColumnFrame> columnFrames, IntConsumer setValue, IntConsumer setStyle) {
        IntConsumer consumer = setValue.andThen(setStyle);
        IntStream.range(0, columnFrames.size()).forEach(consumer);
    }


    private List<Object[]> convertMultiModel(List<?>[] modelLists) {
        Arrays.sort(modelLists, (o1, o2) -> o2.size() - o1.size());
        int maxModelSize = modelLists[0].size();
        int size = modelLists.length;
        List<Object[]> convertModel = new ArrayList<>(maxModelSize);
        IntStream.range(0, maxModelSize).forEach(i -> convertModel.add(new Object[size]));
        for (int i = 0; i < size; i++) {
            List<?> objects = modelLists[i];
            int dataSize = objects.size();
            while (--dataSize >= 0) {
                convertModel.get(dataSize)[i] = objects.get(dataSize);
            }
        }
        return convertModel;
    }

    private List<ColumnFrame> convertImportColumnFrames(List<ColumnFrame> columnFrames) {
        List<ColumnFrame> newList = new ArrayList<>(columnFrames.size());
        int searchRowRange = Math.min(this.sheet.getLastRowNum(), 100);
        for (int i = 0; i < searchRowRange; i++) {
            Stream.of(this.sheet.getRow(i))
                    .filter(Objects::nonNull)
                    .flatMap(row -> StreamSupport.stream(row.spliterator(), false))
                    .filter(Objects::nonNull)
                    .filter(cell -> cell.getCellType() == CellType.STRING)
                    .forEach(cell -> add(columnFrames, newList, cell));
        }
        return newList;
    }

    private void add(List<ColumnFrame> columnFrames, List<ColumnFrame> newList,Cell cell) {
        for (ColumnFrame frame : columnFrames) {
            if (frame.getImportNameOptions().contains(CellUtils.getValue(cell).toString().trim())) {
                columnFrames.remove(frame);
                newList.add(frame.createImportOnlyColumnFrame(cell.getRowIndex(), cell.getColumnIndex()));
                return;
            }
        }
    }


    private List<ColumnFrame> createMultiColumnFrames(Predicate<String> excludedHeader, List<Object[]> multiModel) {
        return Stream.of(multiModel.get(0))
                .map(o -> ClassModelRegistrator.getOrDefault(o.getClass()))
                .flatMap(classModels -> createColumnFrames(classModels, excludedHeader, ColumnFrame :: setNonSheetStyle))
                .collect(Collectors.toList());
    }

    private List<ColumnFrame> createColumnFrames(Predicate<String> excludedHeader, List<ClassModel> classModels) {
        return createColumnFrames(classModels, excludedHeader, frame -> frame.setSheetStyle(this.sheet))
                .collect(Collectors.toList());
    }

    private Stream<ColumnFrame> createColumnFrames(List<ClassModel> classModels,
                                                   Predicate<String> excludedHeader,
                                                   Consumer<ColumnFrame> sheetStyleConsumer) {
        Consumer<ColumnFrame> consumer = sheetStyleConsumer.andThen(ColumnFrame :: setColumnStyle);

        return classModels.stream()
                .map(classModel -> new ColumnFrame(this._wb, classModel))
                .filter(columnFrame -> excludedHeader
                        .negate()
                        .test(columnFrame.getClassModel().getColumn().value()))
                .map(columnFrame -> {
                    consumer.accept(columnFrame);
                    return columnFrame;
                })
                .sorted();
    }

    private Map<String, Object> createModelMap(List<ColumnFrame> columnFrames, int rowIndex) {
        return columnFrames.stream()
                .filter(ColumnFrame :: isExistColumn)
                .collect(Collectors.toMap(
                                frame -> frame.getClassModel().getFieldName(),
                                frame -> CellUtils.getValue(this.sheet, rowIndex, frame.getHeaderColumnPosition())
                        )
                );
    }
}