package org.celper.processor;

import com.google.auto.service.AutoService;
import org.celper.ExcelModel;
import org.celper.processor.parsers.CellFormatAnnotationParser;
import org.celper.processor.parsers.ColumnAnnotationParser;
import org.celper.processor.parsers.DefaultValueAnnotationParser;
import org.celper.processor.util.ElementUtil;
import org.celper.processor.validators.ColumnAnnotationExistValidator;
import org.celper.processor.validators.FieldTypeValidator;
import org.celper.processor.validators.GetterExistValidator;

import javax.annotation.processing.*;
import javax.lang.model.SourceVersion;
import javax.lang.model.element.Element;
import javax.lang.model.element.TypeElement;
import javax.lang.model.element.VariableElement;
import javax.lang.model.util.ElementFilter;
import java.util.*;
import java.util.concurrent.atomic.AtomicInteger;
import java.util.stream.Collectors;
import java.util.stream.Stream;

@SupportedAnnotationTypes("org.celper.ExcelModel")
@SupportedSourceVersion(SourceVersion.RELEASE_8)
@AutoService(Processor.class)
public class Processor extends AbstractProcessor {

    private ElementUtil elementUtil;
    private List<Validator<VariableElement>> validators;
    private List<Parser> parsers;

    @Override
    public synchronized void init(ProcessingEnvironment processingEnv) {
        super.init(processingEnv);
        this.elementUtil = new ElementUtil(processingEnv);

        this.validators = getUnmodifiableList(
                new ColumnAnnotationExistValidator(),
                new FieldTypeValidator(elementUtil),
                new GetterExistValidator(elementUtil));

        this.parsers = getUnmodifiableList(
                new ColumnAnnotationParser(),
                new CellFormatAnnotationParser(),
                new DefaultValueAnnotationParser());
        // TODO style class parser 추가
    }

    @Override
    public boolean process(Set<? extends TypeElement> annotations, RoundEnvironment roundEnv) {
        /**
         * 흐름
         * Column 붙은 애들 lazy
         * Field 타입 확인
         * Getter 확인 - 문제 계속 메서드를 생성해내야함 어거지로 가능하긴하네
         */
        Set<? extends Element> classSet = roundEnv.getElementsAnnotatedWith(ExcelModel.class);
        for (Element clazz : classSet) {
            process(clazz);
        }
//        AtomicInteger atomicInteger = new AtomicInteger(0);
//        // 검증
//        List<FieldAnnotationMetaData> fields = annotations.stream()
//                .map(roundEnv :: getElementsAnnotatedWith)
//                .flatMap(Collection :: stream)
//                .map(Element :: getEnclosedElements)
//                .map(ElementFilter :: fieldsIn)
//                .flatMap(List :: stream)
//                .filter(this :: valid)
//                .map(field -> {
//                    FieldAnnotationMetaData obj = new FieldAnnotationMetaData();
//                    obj.setField(field);
//                    obj.setGetterMethod(elementUtil.buildExpectedGetterName(field));
//                    obj.setDefinedFieldOrder(atomicInteger.getAndIncrement());
//                    return obj;
//                })
//                .collect(Collectors.toList());
        // 파싱
        return true;
    }

    private ClassMetaData process(Element clazz){
        AtomicInteger atomicInteger = new AtomicInteger(0);
        // 검증
        List<FieldAnnotationMetaData> fields = ElementFilter.fieldsIn(clazz.getEnclosedElements())
                .stream()
                .filter(this :: valid)
                .map(field -> {
                    FieldAnnotationMetaData obj = new FieldAnnotationMetaData();
                    obj.setField(field);
                    obj.setGetterMethod(elementUtil.buildExpectedGetterName(field));
                    obj.setDefinedFieldOrder(atomicInteger.getAndIncrement());
                    return obj;
                })
                .map(this::parse)
                .sorted()
                .collect(Collectors.toList());
        ClassMetaData classMetaData = new ClassMetaData();
        // 추가
        classMetaData.setFieldAnnotationMetaDatas(fields);
        return classMetaData;
    }

    @SuppressWarnings("rawtypes")
    private FieldAnnotationMetaData parse(FieldAnnotationMetaData fieldAnnotationMetaData){
        for (Parser parser : parsers) {
            //noinspection unchecked
            parser.parse(fieldAnnotationMetaData);
        }
        return fieldAnnotationMetaData;
    }

    @SuppressWarnings("rawtypes")
    private boolean valid(VariableElement variableElement){
        for (Validator validator : validators) {
            //noinspection unchecked
            if(!validator.valid(variableElement)) return false;
        }
        return true;
    }

    @SafeVarargs
    private final <T> List<T> getUnmodifiableList(T... t){
        return Collections.unmodifiableList(Arrays.asList(t));
    }

}
