package org.celper.processor.parsers;

import org.celper.processor.ClassMetaData;
import org.celper.processor.Parser;
import org.celper.processor.Validator;
import org.celper.processor.validators.ColumnAnnotationExistValidator;
import org.celper.processor.validators.FieldTypeValidator;

import javax.lang.model.element.Element;
import javax.lang.model.element.VariableElement;

public final class ClassElementParser implements Parser<Element, ClassMetaData> {

    private Validator<VariableElement> columnAnnotationValidator = new ColumnAnnotationExistValidator();

    @Override
    public ClassMetaData parse(Element t) {
        //validator
        //parser
        return new ClassMetaData();
    }
}
