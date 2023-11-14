package org.celper.core;

import org.apache.poi.ss.usermodel.*;
import org.celper.core.structure.ColumnStructure;
import org.celper.core.structure.ImportStructure;
import org.celper.core.structure.Structure;
import org.celper.exception.DataListEmptyException;
import org.celper.exception.NoSuchFieldException;
import org.celper.common.ModelMapperFactory;
import org.modelmapper.ModelMapper;

import java.util.*;
import java.util.function.Consumer;
import java.util.function.IntConsumer;
import java.util.function.Predicate;
import java.util.stream.Collectors;
import java.util.stream.IntStream;
import java.util.stream.Stream;
import java.util.stream.StreamSupport;

public class ExcelSheet {
    private final Workbook _wb;
    private final Sheet sheet;
    private final StructureRegistrator structureRegistrator;

    public ExcelSheet(Workbook workbook, Sheet sheet, StructureRegistrator structureRegistrator) {
        this._wb = workbook;
        this.sheet = sheet;
        this.structureRegistrator = structureRegistrator;
    }

    public String getName() {
        return this.sheet.getSheetName();
    }

    public Sheet getSheet() {
        return this.sheet;
    }

    public <T> void modelToSheet(List<T> model) {
        modelToSheet(header -> false, model);
    }

    public <T> void modelToSheet(Predicate<String> excludedHeader, List<T> model) {
        if (Objects.isNull(model) || model.isEmpty()) {
            throw new DataListEmptyException("data list is empty exception");
        }
        List<ColumnStructure> columnStructures = createColumnStructures(excludedHeader, structureRegistrator.getOrDefault(model.get(0).getClass()));

        int headerRowIndex = 0;
        int dataRow = headerRowIndex + 1;
        int modelSize = model.size();

        headerWrite(columnStructures, headerRowIndex);
        IntStream.rangeClosed(dataRow, modelSize)
                .forEach(rowIndex -> dataWrite(columnStructures, rowIndex, model.get(rowIndex - 1)));
    }

    public void multiModelToSheet(List<?>... modelLists) {
        multiModelToSheet(s -> false, modelLists);
    }

    public void multiModelToSheet(Predicate<String> excludedHeader, List<?>... modelLists) {
        for (List<?> modelList : modelLists) {
            if (Objects.isNull(modelList)) {
                throw new DataListEmptyException("data list is empty exception");
            }
        }

        List<Object[]> multiModel = convertMultiModel(modelLists);
        List<ColumnStructure> multiColumnStructures = createMultiColumnStructures(excludedHeader, multiModel);

        int headerRowIndex = 0;
        int dataRow = headerRowIndex + 1;
        int multiModelSize = multiModel.size();

        headerWrite(multiColumnStructures, headerRowIndex);
        IntStream.rangeClosed(dataRow, multiModelSize)
                .forEach(rowIndex -> dataWrite(multiColumnStructures, rowIndex, multiModel.get(rowIndex - 1)));
    }

    public <T> List<T> sheetToModel(Class<T> clazz) {
        ModelMapper modelMapper = ModelMapperFactory.defaultModelMapper(); // Model Mapper 가져오기

        List<ColumnStructure> columnStructures = createColumnStructures(header -> false, structureRegistrator.getOrDefault(clazz));
        List<ImportStructure> importList = convertImportStructures(columnStructures);

        int startRow = importList.stream()// 병합을 고려해서 startRow 추출 병합 고려할려면 isMerged 만들어야함
                .max(ImportStructure :: compareTo)
                .orElseThrow(() -> new NoSuchFieldException(String.format("'%s' 시트에 매칭된 필드가 없습니다.", this.sheet.getSheetName())))
                .getHeaderRowPosition() + 1;

        return IntStream.rangeClosed(startRow, this.sheet.getLastRowNum())
                .filter(rowIdx -> Objects.nonNull(this.sheet.getRow(rowIdx)))
                .mapToObj(rowIdx -> createModelMap(importList, rowIdx))
                .map(entry -> modelMapper.map(entry, clazz))
                .collect(Collectors.toList());
    }

    private void headerWrite(List<ColumnStructure> columnStructures, int rowIndex) {
        Row headerRow = CellUtils.createRow(this.sheet, rowIndex, columnStructures.size());
        IntConsumer setHeader = colIdx -> CellUtils.setValue(headerRow.getCell(colIdx), columnStructures.get(colIdx).getStructure().getColumn().value());
        IntConsumer setStyle = colIdx -> headerRow.getCell(colIdx).setCellStyle(columnStructures.get(colIdx).getHeaderAreaCellStyle());
        write(columnStructures, setHeader, setStyle);
    }

    private void dataWrite(List<ColumnStructure> columnStructures, int rowIndex, Object[] model) {
        Stream.of(model).forEach(o -> dataWrite(columnStructures, rowIndex, o));
    }

    private void dataWrite(List<ColumnStructure> columnStructures, int rowIndex, Object o) {
        Row row = CellUtils.createRow(this.sheet, rowIndex, columnStructures.size());
        IntConsumer setValue = colIdx -> CellUtils.setValue(columnStructures.get(colIdx), row.getCell(colIdx), o);
        IntConsumer setStyle = colIdx -> row.getCell(colIdx).setCellStyle(columnStructures.get(colIdx).getDataAreaCellStyle());
        write(columnStructures, setValue, setStyle);
    }

    private void write(List<ColumnStructure> columnStructures, IntConsumer setValue, IntConsumer setStyle) {
        IntConsumer consumer = setValue.andThen(setStyle);
        IntStream.range(0, columnStructures.size()).forEach(consumer);
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

    private List<ImportStructure> convertImportStructures(List<ColumnStructure> columnStructures) {
        List<ImportStructure> newList = new ArrayList<>(columnStructures.size());
        int searchRowRange = Math.min(this.sheet.getLastRowNum(), 100);
        for (int i = 0; i < searchRowRange; i++) {
            Stream.of(this.sheet.getRow(i))
                    .filter(Objects :: nonNull)
                    .flatMap(row -> StreamSupport.stream(row.spliterator(), false))
                    .filter(Objects :: nonNull)
                    .filter(cell -> cell.getCellType() == CellType.STRING)
                    .forEach(cell -> add(columnStructures, newList, cell));
        }
        return newList;
    }

    private void add(List<ColumnStructure> columnStructures, List<ImportStructure> newList, Cell cell) {
        for (ColumnStructure structure : columnStructures) {
            if (structure.getImportNameOptions().contains(CellUtils.getValue(cell).toString().trim().intern())) {
                ImportStructure importStructure = new ImportStructure(structure, cell.getRowIndex(), cell.getColumnIndex());
                newList.add(importStructure);
                columnStructures.remove(structure);
                return;
            }
        }
    }

    private List<ColumnStructure> createMultiColumnStructures(Predicate<String> excludedHeader, List<Object[]> multiModel) {
        return Stream.of(multiModel.get(0))
                .map(o -> structureRegistrator.getOrDefault(o.getClass()))
                .flatMap(structure -> createColumnStructures(structure, excludedHeader, ColumnStructure :: setNonSheetStyle))
                .collect(Collectors.toList());
    }

    private List<ColumnStructure> createColumnStructures(Predicate<String> excludedHeader, List<Structure> structures) {
        return createColumnStructures(structures, excludedHeader, structure -> structure.setSheetStyle(this.sheet))
                .collect(Collectors.toList());
    }

    private Stream<ColumnStructure> createColumnStructures(List<Structure> structures,
                                                           Predicate<String> excludedHeader,
                                                           Consumer<ColumnStructure> sheetStyleConsumer) {
        Consumer<ColumnStructure> consumer = sheetStyleConsumer.andThen(ColumnStructure :: setColumnStyle);
        return structures.stream()
                .map(structure -> new ColumnStructure(this._wb, structure))
                .filter(columnStructure -> excludedHeader
                        .negate()
                        .test(columnStructure.getStructure().getColumn().value()))
                .map(columnStructure -> {
                    consumer.accept(columnStructure);
                    return columnStructure;
                })
                .sorted();
    }

    private Map<String, Object> createModelMap(List<ImportStructure> importStructure, int rowIndex) {
        return importStructure.stream()
                .filter(ImportStructure :: existPosition)
                .collect(Collectors.toMap(
                                structure -> structure.getStructure().getFieldName(),
                                structure -> CellUtils.getValue(this.sheet, rowIndex, structure.getHeaderColumnPosition())
                        )
                );
    }
}