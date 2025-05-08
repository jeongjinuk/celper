package org.celper.processor;

import com.google.auto.service.AutoService;
import com.squareup.javapoet.JavaFile;
import org.celper.ExcelModel;
import org.celper.processor.generator.ClassGenerator;
import org.celper.processor.meta.ClassMetaData;
import org.celper.processor.parser.ClassParser;
import org.celper.processor.util.ElementUtil;

import javax.annotation.processing.*;
import javax.lang.model.SourceVersion;
import javax.lang.model.element.TypeElement;
import javax.tools.Diagnostic;
import java.io.BufferedWriter;
import java.io.IOException;
import java.io.OutputStreamWriter;
import java.io.Writer;
import java.nio.charset.StandardCharsets;
import java.util.List;
import java.util.Optional;
import java.util.Set;
import java.util.function.Function;
import java.util.stream.Collectors;

@SupportedAnnotationTypes({
        "org.celper.ExcelModel",
        "org.celper.SheetLayout",
        "org.celper.SheetStyle",
        "org.celper.Column",
        "org.celper.ColumnStyle",
        "org.celper.DefaultValue",
        "org.celper.CellFormat",
        "org.celper.CSVConfig"
})
@SupportedSourceVersion(SourceVersion.RELEASE_8)
@AutoService(Processor.class)
public class ExcelProcessor extends AbstractProcessor {

    private ElementUtil elementUtil;
    private ClassParser classParser;
    private ClassGenerator classGenerator;
    private boolean processed = false;

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
        this.classGenerator = new ClassGenerator(elementUtil);
    }

    @Override
    public boolean process(Set<? extends TypeElement> annotations, RoundEnvironment roundEnv) {
        if (!processed){
            processed = true;
            List<ClassMetaData> classMetaDataList = roundEnv.getElementsAnnotatedWith(ExcelModel.class)
                    .stream()
                    .map(element -> (TypeElement) element)
                    .map(classParser :: parse)
                    .filter(Optional :: isPresent)
                    .map(Optional :: get)
                    .collect(Collectors.toList());
            elementUtil.log(Diagnostic.Kind.NOTE, READY_TO_GENERATED_CLASSES_LIST_LOG_MSG.apply(classMetaDataList));
            for (ClassMetaData classMetaData : classMetaDataList) {
                write(classGenerator.generate(classMetaData));
            }
        }
        return processed;
    }


    private void write(JavaFile javaFile){
        String path = javaFile.packageName.length() == 0 ? javaFile.typeSpec.name : javaFile.packageName + "." + javaFile.typeSpec.name;
        try(Writer writer = getWriter(path)){
            javaFile.writeTo(writer);
        }catch (IOException e){
            e.printStackTrace();
        }
    }

    private Writer getWriter(String path) throws IOException {
        return new BufferedWriter(
                new OutputStreamWriter(processingEnv.getFiler().createSourceFile(path).openOutputStream(),
                StandardCharsets.UTF_8)
        );
    }

}
