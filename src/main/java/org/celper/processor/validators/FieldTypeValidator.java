package org.celper.processor.validators;

import org.celper.processor.Validator;

import javax.annotation.processing.ProcessingEnvironment;
import javax.lang.model.element.TypeElement;
import javax.lang.model.element.VariableElement;
import javax.lang.model.type.PrimitiveType;
import javax.lang.model.type.TypeMirror;
import java.math.BigDecimal;
import java.math.BigInteger;
import java.time.LocalDate;
import java.time.LocalDateTime;
import java.time.LocalTime;

public class FieldTypeValidator implements Validator<VariableElement> {
    /**
     * TODO javaDoc 추가
     *
     * 전환 방식
     * 모든 primitive 검사 -> 통과시 -> boxed 변환 -> SupportedType 일때 true
     * TypeKind.isPrimitive() -> true ? primitiveType : typeMirror
     * switch
     * processingEnv.getElementUtils().getTypeElement(String.class.getCanonicalName()).asType(); // 여기 존재하는 타입만 지원
     * processingEnv.getTypeUtils().boxedClass((PrimitiveType) typeMirror) 박싱해서 내 타입이 맞는지 확인
     * processingEnv.getTypeUtils().isSameType(element.asType(),string) // 확인
     *
     *  @일반형
     *  java.lang.Integer
     *  java.lang.Float
     *  java.lang.Boolean
     *  java.lang.Double
     *  java.lang.Long
     *
     *  @문자형
     *  java.lang.String
     *
     *  @특수형
     *  java.time.LocalDateTime
     *  java.time.LocalDate
     *  java.time.LocalTime
     */
    private final ProcessingEnvironment processingEnv;

    public FieldTypeValidator(ProcessingEnvironment processingEnv) {
        this.processingEnv = processingEnv;
    }

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

        public TypeMirror getTypeMirror(ProcessingEnvironment processingEnvironment) {
            return processingEnvironment.getElementUtils().getTypeElement(clazz.getCanonicalName()).asType();
        }
    }

    @Override
    public boolean valid(VariableElement variableElement) {
        TypeMirror typeMirror = variableElement.asType();
        typeMirror = typeMirror.getKind().isPrimitive() ?
                processingEnv.getTypeUtils().boxedClass((PrimitiveType) variableElement).asType() : typeMirror;
        for (SupportType supportType : SupportType.values()) {
            if (processingEnv.getTypeUtils().isSameType(typeMirror, supportType.getTypeMirror(processingEnv)))
                return true;
        }
        throw new UnsupportedOperationException(
                String.format("The field '%s' in class '%s' is of unsupported type '%s.'",
                        variableElement.getSimpleName(),
                        ((TypeElement) variableElement.getEnclosingElement()).getQualifiedName()));
    }
}
