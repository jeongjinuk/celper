package org.celper.processor.validators;

import org.celper.processor.Validator;
import org.celper.processor.util.ElementUtil;

import javax.lang.model.element.VariableElement;
import javax.lang.model.util.ElementFilter;

public class GetterExistValidator implements Validator<VariableElement> {
    /**
     * TODO javaDoc 추가
     */

    private final ElementUtil util;

    public GetterExistValidator(ElementUtil util) {
        this.util = util;
    }

    @Override
    public boolean valid(VariableElement field) {
        ElementFilter.methodsIn(field.getEnclosingElement().getEnclosedElements())
                .stream()
                .filter(method -> util.isGetterMethodForField(method, field))
                .filter(util :: isNonParameter)
                .filter(method -> util.isTypeMatch(method, field))
                .findAny()
                .orElseThrow(() -> new IllegalStateException(
                        String.format("The getter method for the field '%s' was not found in the class '%s'. " +
                                        "Please ensure that the getter method follows the naming convention '%s'.",
                                field.getSimpleName(),
                                field.getEnclosingElement().getSimpleName(),
                                util.buildExpectedGetterName(field))));
        return true;
    }
}