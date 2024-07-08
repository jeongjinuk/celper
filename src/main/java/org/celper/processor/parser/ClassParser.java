package org.celper.processor.parser;

import org.celper.Column;
import org.celper.processor.meta.ClassMetaData;
import org.celper.processor.meta.FieldAnnotationMetaData;
import org.celper.processor.util.ElementUtil;
import org.celper.processor.validator.Validator;
import org.celper.processor.validator.ValidatorFactory;

import javax.lang.model.element.Element;
import javax.lang.model.element.TypeElement;
import javax.lang.model.element.VariableElement;
import javax.tools.Diagnostic;
import java.util.List;
import java.util.Optional;
import java.util.concurrent.atomic.AtomicInteger;
import java.util.stream.Collectors;

public class ClassParser {
    private final ElementUtil elementUtil;
    private final List<Validator<VariableElement>> validators;

    public ClassParser(ElementUtil elementUtil) {
        this.elementUtil = elementUtil;
        this.validators = ValidatorFactory.getValidators(elementUtil);
    }

    public Optional<ClassMetaData> parse(TypeElement clazz) {
        List<VariableElement> fields = elementUtil.getFieldsWithAnnotation(clazz, Column.class);

        if (fields.isEmpty()){
            elementUtil.log(Diagnostic.Kind.WARNING,
                    String.format("The class '%s' annotated with @ExcelModel does not contain any fields annotated with @Column. " +
                                    "Please ensure that at least one field is annotated with @Column.",
                            clazz.getSimpleName()));
            return Optional.empty();
        }
        validateFields(fields); // 예외
        // 추가
        List<FieldAnnotationMetaData> fieldAnnotationMetaData = parseFields(fields);
        return Optional.of(createClassMetaData(clazz, fieldAnnotationMetaData));
    }

    private void validateFields(List<VariableElement> fields){
        fields.forEach(this::validateField);
    }
    private void validateField(VariableElement field){
        for (Validator<VariableElement> validator : validators) {
            validator.valid(field);
        }
    }
    private ClassMetaData createClassMetaData(TypeElement clazz, List<FieldAnnotationMetaData> fieldAnnotationMetaData){
        ClassMetaData classMetaData = new ClassMetaData(clazz, fieldAnnotationMetaData);
        AnnotationHandler.applyHandlersIfPresent(ClassAnnotationHandler.values(), clazz, classMetaData);
        return classMetaData;
    }

    private List<FieldAnnotationMetaData> parseFields(List<VariableElement> fields){
        AtomicInteger atomicInteger = new AtomicInteger(0);
        return fields
                .stream()
                .map(field -> createFieldMetaData(field, atomicInteger.getAndIncrement()))
                .sorted()
                .collect(Collectors.toList());
    }

    private FieldAnnotationMetaData createFieldMetaData(VariableElement field, int defindOrder){
        FieldAnnotationMetaData metaData = new FieldAnnotationMetaData();
        metaData.setField(field);
        metaData.setDefinedFieldOrder(defindOrder);
        metaData.setGetterMethod(elementUtil.generateGetterName(field));
        AnnotationHandler.applyHandlersIfPresent(FieldAnnotationHandler.values(), field, metaData);
        return metaData;
    }


}
