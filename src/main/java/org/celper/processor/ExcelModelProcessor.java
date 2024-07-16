//package org.celper.processor;
//
//import com.google.auto.service.AutoService;
//import lombok.SneakyThrows;
//import org.celper.ExcelModel;
//
//import javax.annotation.processing.*;
//import javax.lang.model.SourceVersion;
//import javax.lang.model.element.Element;
//import javax.lang.model.element.TypeElement;
//import javax.tools.JavaFileObject;
//import java.io.IOException;
//import java.io.Writer;
//import java.util.HashSet;
//import java.util.Set;
//
//@SupportedAnnotationTypes("org.celper.ExcelModel")
//@SupportedSourceVersion(SourceVersion.RELEASE_8)
//@AutoService(Processor.class)
//public class ExcelModelProcessor extends AbstractProcessor {
//
//    private Set<String> modelImplementations = new HashSet<>();
//
//    @Override
//    public boolean process(Set<? extends TypeElement> annotations, RoundEnvironment roundEnv) {
//        for (Element element : roundEnv.getElementsAnnotatedWith(ExcelModel.class)) {
//            if (element instanceof TypeElement) {
//                TypeElement typeElement = (TypeElement) element;
//                generateModelClass(typeElement);
//            }
//        }
//        return false;
//    }
//
//    @SneakyThrows
//    private void generateModelClass(TypeElement typeElement) {
//        String className = typeElement.getSimpleName().toString();
//        String packageName = processingEnv.getElementUtils().getPackageOf(typeElement).getQualifiedName().toString();
//        String newClassName = className + "ExcelModel";
//        String fullClassName = packageName + "." + newClassName;
//
//        // Generate the new class in the client's project
//        Filer filer = processingEnv.getFiler();
//        JavaFileObject javaFileObject = filer.createSourceFile(fullClassName, typeElement);
//        try (Writer writer = javaFileObject.openWriter()) {
//            writer.write("package " + packageName + ";\n\n");
//            writer.write("//테스트용 \n\n");
//            writer.write("import org.celper.Model;\n");
//            writer.write("public class " + newClassName + " implements Model<" + className + "> {\n");
//            writer.write("    @Override\n");
//            writer.write("    public void print(" + className + " instance) {\n");
//
//            typeElement.getEnclosedElements().stream()
//                    .filter(e -> e.getKind().isField())
//                    .forEach(field -> {
//                        String fieldName = field.getSimpleName().toString();
//                        String getterName = "get" + fieldName.substring(0, 1).toUpperCase() + fieldName.substring(1) + "()";
//                        try {
//                            writer.write("        System.out.println(instance." + getterName + ");\n");
//                        } catch (IOException e) {
//                            processingEnv.getMessager().printMessage(javax.tools.Diagnostic.Kind.ERROR, e.toString());
//                        }
//                    });
//            writer.write("    }\n");
//            writer.write("}\n");
//        }
//
//        modelImplementations.add(fullClassName);
//    }
//}
