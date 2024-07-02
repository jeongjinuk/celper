package org.celper.processor.validators;

import org.celper.processor.Validator;

import javax.annotation.processing.ProcessingEnvironment;
import javax.lang.model.element.ExecutableElement;
import javax.lang.model.element.VariableElement;
import javax.lang.model.util.ElementFilter;
import java.util.function.Function;

public class GetterExistValidator implements Validator<VariableElement> {
    /**
     * TODO javaDoc 추가
     * processing Env 주입여부 Error message 때문에
     */
    private static final String GET = "get";
    private static final Function<String, StringBuilder> buildExpectedGetterName = fieldName -> new StringBuilder(GET)
                    .append(Character.toUpperCase(fieldName.charAt(0)))
                    .append(fieldName.substring(1));

    private final ProcessingEnvironment processingEnv;

    public GetterExistValidator(ProcessingEnvironment processingEnv) {
        this.processingEnv = processingEnv;
    }

    @Override
    public boolean valid(VariableElement field) {
        ElementFilter.methodsIn(field.getEnclosingElement().getEnclosedElements())
                .stream()
                .filter(method -> isGetterMethodForField(method, field))
                .filter(this :: isNonParameter)
                .filter(method -> isTypeMatch(method, field))
                .findAny()
                .orElseThrow(() -> new IllegalStateException(
                        String.format("The getter method for the field '%s' was not found in the class '%s'. " +
                                        "Please ensure that the getter method follows the naming convention '%s'.",
                                field.getSimpleName(),
                                field.getEnclosingElement().getSimpleName(),
                                buildExpectedGetterName.apply(field.getSimpleName().toString()))));
        return true;
    }

    private boolean isGetterMethodForField(ExecutableElement method, VariableElement field) {
        String fieldName = field.getSimpleName().toString();
        return method.getSimpleName().contentEquals(buildExpectedGetterName.apply(fieldName));
    }

    private boolean isNonParameter(ExecutableElement method) {
        return method.getParameters().size() == 0;
    }

    private boolean isTypeMatch(ExecutableElement method, VariableElement field){
        return processingEnv.getTypeUtils().isSameType(method.getReturnType(), field.asType());
    }
}
