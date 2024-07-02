package org.celper.processor.validators;

import org.celper.Column;
import org.celper.processor.Validator;

import javax.lang.model.element.VariableElement;
import java.util.Objects;

public final class ColumnAnnotationExistValidator implements Validator<VariableElement> {
    @Override
    public boolean valid(VariableElement variableElement) {
        return Objects.nonNull(variableElement.getAnnotation(Column.class));
    }
}
