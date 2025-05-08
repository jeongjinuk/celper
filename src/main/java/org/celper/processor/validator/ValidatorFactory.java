package org.celper.processor.validator;

import org.celper.processor.util.ElementUtil;

import javax.lang.model.element.VariableElement;
import java.util.Arrays;
import java.util.Collections;
import java.util.List;

public final class ValidatorFactory {

    private ValidatorFactory() {
    }

    public static List<Validator<VariableElement>> getValidators(ElementUtil elementUtil) {
        return Collections.unmodifiableList(
                Arrays.asList(
                        new FieldTypeValidator(elementUtil),
                        new GetterExistValidator(elementUtil)
                )
        );
    }
}
