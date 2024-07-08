package org.celper.processor;

import com.google.auto.service.AutoService;
import org.celper.ExcelModel;
import org.celper.processor.meta.ClassMetaData;
import org.celper.processor.parser.ClassParser;
import org.celper.processor.util.ElementUtil;

import javax.annotation.processing.*;
import javax.lang.model.SourceVersion;
import javax.lang.model.element.TypeElement;
import javax.tools.Diagnostic;
import java.util.List;
import java.util.Optional;
import java.util.Set;
import java.util.function.Function;
import java.util.stream.Collectors;

@SupportedAnnotationTypes("org.celper.ExcelModel")
@SupportedSourceVersion(SourceVersion.RELEASE_8)
@AutoService(Processor.class)
public class Processor extends AbstractProcessor {

    private ElementUtil elementUtil;
    private ClassParser classParser;

    private static final Function<List<ClassMetaData>, String> READY_TO_GENERATED_CLASSES_LIST_LOG_MSG =
            classMetaData -> String.format("The following classes are ready to be generated: [%s]",
                    classMetaData.stream()
                            .map(ClassMetaData :: getClazz)
                            .map(TypeElement :: getQualifiedName)
                            .collect(Collectors.joining(", ")));


    @Override
    public synchronized void init(ProcessingEnvironment processingEnv) {
        super.init(processingEnv);
        this.elementUtil = new ElementUtil(processingEnv);
        this.classParser = new ClassParser(elementUtil);
    }

    @Override
    public boolean process(Set<? extends TypeElement> annotations, RoundEnvironment roundEnv) {
        List<ClassMetaData> classMetaDataList = roundEnv.getElementsAnnotatedWith(ExcelModel.class)
                .stream()
                .map(element -> (TypeElement) element)
                .map(classParser :: parse)
                .filter(Optional :: isPresent)
                .map(Optional :: get)
                .collect(Collectors.toList());
        elementUtil.log(Diagnostic.Kind.NOTE, READY_TO_GENERATED_CLASSES_LIST_LOG_MSG.apply(classMetaDataList));
        // classMetaData -> generated Class
        return true;
    }

}
