package org.celper.processor.validator;

import org.celper.processor.util.ElementUtil;

import javax.lang.model.element.VariableElement;
import javax.tools.Diagnostic;

public final class GetterExistValidator implements Validator<VariableElement> {
    /**
     * TODO javaDoc 추가
     */

    private final ElementUtil util;
    private static final String logFormat  =
            "The getter method for the field '%s' was not found in the class '%s'. " +
            "Please ensure that the getter method follows the naming convention '%s'.";

    public GetterExistValidator(ElementUtil util) {
        this.util = util;
    }

    @Override
    public void valid(VariableElement field) {
        if(!util.hasMatchingGetter(field)){
            String msg = String.format(logFormat,
                    field.getSimpleName(),
                    field.getEnclosingElement().getSimpleName(),
                    util.generateGetterName(field));
            util.log(Diagnostic.Kind.ERROR, msg);
        }
    }
}