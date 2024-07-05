package org.celper.processor.util;

import javax.annotation.processing.Messager;
import javax.annotation.processing.ProcessingEnvironment;
import javax.lang.model.element.Element;
import javax.lang.model.element.ExecutableElement;
import javax.lang.model.element.VariableElement;
import javax.lang.model.type.PrimitiveType;
import javax.lang.model.type.TypeMirror;
import javax.lang.model.util.ElementFilter;
import javax.lang.model.util.Elements;
import javax.lang.model.util.Types;
import javax.tools.Diagnostic;
import java.lang.annotation.Annotation;
import java.math.BigDecimal;
import java.math.BigInteger;
import java.time.LocalDate;
import java.time.LocalDateTime;
import java.time.LocalTime;
import java.util.List;
import java.util.Objects;
import java.util.function.Function;
import java.util.stream.Collectors;

public final class ElementUtil {

    private static final String GET = "get";

    private static final Function<String, StringBuilder> buildExpectedGetterName = fieldName -> new StringBuilder(GET)
            .append(Character.toUpperCase(fieldName.charAt(0)))
            .append(fieldName.substring(1));

    private enum SupportType {
        INTEGER(Integer.class),
        FLOAT(Float.class),
        BOOLEAN(Boolean.class),
        DOUBLE(Double.class),
        LONG(Long.class),
        STRING(String.class),
        LOCAL_DATE_TIME(LocalDateTime.class),
        LOCAL_DATE(LocalDate.class),
        LOCAL_TIME(LocalTime.class),
        BIG_DECIMAL(BigDecimal.class),
        BIG_INTEGER(BigInteger.class);

        private Class clazz;

        SupportType(Class clazz) {
            this.clazz = clazz;
        }
    }

    private final Types typeUtils;
    private final Elements elementUtils;
    private final Messager messager;

    public ElementUtil(ProcessingEnvironment processingEnv) {
        this.typeUtils = processingEnv.getTypeUtils();
        this.elementUtils = processingEnv.getElementUtils();
        this.messager = processingEnv.getMessager();
    }

    public List<VariableElement> getFieldsWithAnnotation(Element clazz, Class<? extends Annotation> annotation){
        return ElementFilter.fieldsIn(clazz.getEnclosedElements())
                .stream()
                .filter(field -> Objects.nonNull(field.getAnnotation(annotation)))
                .collect(Collectors.toList());
    }

    public String generateGetterName(VariableElement field) {
        return buildExpectedGetterName.apply(field.getSimpleName().toString()).toString();
    }

    public boolean isFieldTypeSupported(VariableElement field) {
        for (SupportType supportType : SupportType.values()) {
            TypeMirror typeMirror = elementUtils.getTypeElement(supportType.clazz.getCanonicalName()).asType();
            if (typeUtils.isSameType(typeMirror, convertToBoxedType(field)))
                return true;
        }
        return false;
    }

    public boolean hasMatchingGetter(VariableElement field) {
        return ElementFilter.methodsIn(field.getEnclosingElement().getEnclosedElements())
                .stream()
                .anyMatch(method -> matchesGetterSignature(method, field));
    }
    public void log(Diagnostic.Kind kind, String msg){
        messager.printMessage(kind, msg);
    }


    private TypeMirror convertToBoxedType(VariableElement field) {
        if (field.asType().getKind().isPrimitive()){
            return  typeUtils.boxedClass((PrimitiveType) field).asType();
        }
        return field.asType();
    }

    private boolean matchesGetterSignature(ExecutableElement method, VariableElement field) {
        String fieldName = field.getSimpleName().toString();
        return method.getSimpleName().contentEquals(buildExpectedGetterName.apply(fieldName)) &&
                (method.getParameters().isEmpty() || method.getParameters().size() == 0) &&
                typeUtils.isSameType(method.getReturnType(), field.asType());
    }
}
