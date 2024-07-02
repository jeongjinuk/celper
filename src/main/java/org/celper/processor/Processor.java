package org.celper.processor;

import com.google.auto.service.AutoService;
import org.celper.processor.validators.FieldTypeValidator;
import org.celper.processor.validators.GetterExistValidator;

import javax.annotation.processing.AbstractProcessor;
import javax.annotation.processing.RoundEnvironment;
import javax.annotation.processing.SupportedAnnotationTypes;
import javax.annotation.processing.SupportedSourceVersion;
import javax.lang.model.SourceVersion;
import javax.lang.model.element.Element;
import javax.lang.model.element.TypeElement;
import javax.lang.model.element.VariableElement;
import javax.lang.model.util.ElementFilter;
import java.util.*;

@SupportedAnnotationTypes("org.celper.ExcelModel")
@SupportedSourceVersion(SourceVersion.RELEASE_8)
@AutoService(Processor.class)
public class Processor extends AbstractProcessor {

    private final List<Validator<VariableElement>> validators;

    public Processor() {
        validators = getUnmodifiableList(
                new FieldTypeValidator(processingEnv),
                new GetterExistValidator(processingEnv));
    }

    private <T> List<T> getUnmodifiableList(T... t){
        return Collections.unmodifiableList(Arrays.asList(t));
    }

    @Override
    public boolean process(Set<? extends TypeElement> annotations, RoundEnvironment roundEnv) {
        /**
         * 흐름
         * Column 붙은 애들 lazy
         * Field 타입 확인
         * Getter 확인 - 문제 계속 메서드를 생성해내야함 어거지로 가능하긴하네
         */
        annotations.stream()
                .map(roundEnv::getElementsAnnotatedWith)
                .flatMap(Collection::stream)
                .map(Element::getEnclosedElements)
                .map(ElementFilter::fieldsIn)
                .flatMap(List::stream)
                .filter(this::valid);

        return true;
    }

    private boolean valid(VariableElement variableElement){
        for (Validator validator : validators) {
            if(!validator.valid(variableElement)) return false;
        }
        return true;
    }
}
