package org.celper.processor.parsers;

import org.celper.processor.FieldAnnotationMetaData;
import org.celper.processor.Parser;

import javax.lang.model.element.Element;
import javax.lang.model.element.ExecutableElement;
import javax.lang.model.util.ElementFilter;
import java.util.List;

class ColumnAnnotationParser implements Parser<Element, List<FieldAnnotationMetaData>> {



    public ColumnAnnotationParser() {
    }

    @Override
    public List<FieldAnnotationMetaData> parse(Element element) {
//        ElementFilter.fieldsIn(element.getEnclosedElements())
//                .stream()
//                .filter();

        List<ExecutableElement> executableElements = ElementFilter.methodsIn(element.getEnclosedElements());

        return null;
    }



}
