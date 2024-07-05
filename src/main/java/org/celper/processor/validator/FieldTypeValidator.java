package org.celper.processor.validator;

import org.celper.processor.util.ElementUtil;

import javax.lang.model.element.TypeElement;
import javax.lang.model.element.VariableElement;
import javax.tools.Diagnostic;

public final class FieldTypeValidator implements Validator<VariableElement> {
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
    private final ElementUtil util;
    private static final String logFormat  = "The field '%s' in class '%s' is of unsupported type '%s.'";
    public FieldTypeValidator(ElementUtil util) {
        this.util = util;
    }

    @Override
    public void valid(VariableElement variableElement) {
        if (!util.isFieldTypeSupported(variableElement)) {
            String msg = String.format(logFormat,
                    variableElement.getSimpleName(),
                    ((TypeElement) variableElement.getEnclosingElement()).getQualifiedName(),
                    variableElement.asType());
            util.log(Diagnostic.Kind.ERROR, msg);
        }
    }
}