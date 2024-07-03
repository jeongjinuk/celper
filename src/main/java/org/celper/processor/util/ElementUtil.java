package org.celper.processor.util;

import javax.annotation.processing.Messager;
import javax.annotation.processing.ProcessingEnvironment;
import javax.lang.model.element.Element;
import javax.lang.model.element.ExecutableElement;
import javax.lang.model.element.VariableElement;
import javax.lang.model.type.PrimitiveType;
import javax.lang.model.type.TypeMirror;
import javax.lang.model.util.Elements;
import javax.lang.model.util.Types;
import javax.tools.Diagnostic;
import java.math.BigDecimal;
import java.math.BigInteger;
import java.time.LocalDate;
import java.time.LocalDateTime;
import java.time.LocalTime;
import java.util.function.Function;

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

    public String buildExpectedGetterName(VariableElement field){
        return buildExpectedGetterName.apply(field.getSimpleName().toString()).toString();
    }

    public TypeMirror toBoxedType(VariableElement field){
        return isPrimitive(field) ? typeUtils.boxedClass((PrimitiveType) field).asType() : field.asType();
    }

    public TypeMirror getTypeMirror(Class<?> clazz){
        return elementUtils.getTypeElement(clazz.getCanonicalName()).asType();
    }

    public boolean isSupportedFieldType(VariableElement field){
        for (SupportType value : SupportType.values()) {
            if (typeUtils.isSameType(getTypeMirror(value.clazz), toBoxedType(field)))
                return true;

        }
        return false;
    }

    public boolean isPrimitive(Element element){
        return element.asType().getKind().isPrimitive();
    }

    public boolean isGetterMethodForField(ExecutableElement method, VariableElement field) {
        String fieldName = field.getSimpleName().toString();
        return method.getSimpleName().contentEquals(buildExpectedGetterName.apply(fieldName));
    }

    public boolean isNonParameter(ExecutableElement method) {
        return method.getParameters().size() == 0;
    }

    public boolean isTypeMatch(ExecutableElement method, VariableElement field){
        return typeUtils.isSameType(method.getReturnType(), field.asType());
    }

}
