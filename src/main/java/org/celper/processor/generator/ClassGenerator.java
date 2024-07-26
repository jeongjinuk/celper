package org.celper.processor.generator;

import com.squareup.javapoet.*;
import org.celper.MetaData;
import org.celper.processor.meta.ClassMetaData;
import org.celper.processor.meta.FieldAnnotationMetaData;
import org.celper.processor.util.ElementUtil;
import org.celper.register.RegisterManager;

import javax.lang.model.element.ExecutableElement;
import javax.lang.model.element.Modifier;
import javax.lang.model.type.TypeMirror;
import java.util.*;
import java.util.stream.Collectors;

/**
 * @FQNC FQNC = Fully Qualified Class Name(클래스가 속한 패키지 이름을 포함하고, 다른 클래스와 충돌하지 않음.)
 * 예) java.util.List<java.lang.Object> == List<Object>
 * @필요_항목 1. DTO 클래스의 FQNC
 * 2. MetaData interface FQNC
 * 3. Generate Class Name
 * 4. Generate Class FQNC
 * @대응_명칭 1. ClassName targetClassFQNC = ClassName.get(classMetaData.getClazz());
 * 2. ClassName interfaceFQNC = ClassName.get(MetaData.class);
 * 3. String generatedClassName = targetClassFQNC.simpleName() + SUFFIX;
 * 4. ClassName generatedClassFQNC = ClassName.get(targetClassFQNC.packageName(), generatedClassName);
 * @가안 package [ targetClassFQNC.getPackageName() ];
 * @import_javapoet에_의해_생성 import org.celper.core.style.*;
 * import org.celper.register.RegisterManager;
 * import java.util.*;
 * import java.util.function.*;
 * @추가_설명 (null 또는 오브젝트 형태로 만든 이유는 기본값을 기본형으로 처리하기 힘들다. 그리고 optional도 고려를 했는데, 이부분에서는 더 복잡해지고 차후 추가될 로직에 대해 알기 어려워서 null로 처리)
 * public final class [ generatedClassName ] implements [ interfaceFQNC ]<[ targetClassFQNC ]> {
 * private static final [ interfaceFQNC ]<[ targetClassFQNC ]> INSTANCE = new [ generatedClassFQNC ]();
 * private final SheetStyleConfigurer SheetStyle = [ null or SheetStyleConfig ];
 * private final List<CellStyleConfigurer> headerStyleConfigList = Collections.unmodifiableList(Arrays.asList( [ null or CellStyleConfig ] ));
 * private final List<CellStyleConfigurer> dataStyleConfigList = Collections.unmodifiableList(Arrays.asList( [ null or CellStyleConfig ] ));
 * private final List<String> columnNameList = Collections.unmodifiableList(Arrays.asList( [ null or String Type ] ));
 * private final List<String> defaultValueList = Collections.unmodifiableList(Arrays.asList( [ null or String Type ] ));
 * private final List<String> cellFormatList = Collections.unmodifiableList(Arrays.asList( [ null or String Type ] ));
 * private final List<Function<DTO, Object>> getterFunctionList = Collections.unmodifiableList(Arrays.asList( [ null or targetClassFQNC :: getterMethodName ] ));
 * <p>
 * static {
 * RegisterManager.put([ targetClassFQNC ], INSTANCE);
 * }
 * @MetaData_interface_Overrided_Methods... }
 */
public class ClassGenerator implements Generator<ClassMetaData, JavaFile> {

    private static final String SUFFIX = "_MetaData";

    private final ElementUtil elementUtil;

    public ClassGenerator(ElementUtil elementUtil) {
        this.elementUtil = elementUtil;
    }

    @Override
    public JavaFile generate(ClassMetaData classMetaData) {
        ClassName targetClassFQNC = ClassName.get(classMetaData.getClazz()); // DTO FQNC
        ClassName interfaceFQNC = ClassName.get(MetaData.class); // interface MetaData FQNC
        String generatedClassName = targetClassFQNC.simpleName() + SUFFIX; // generate Class Name
        ClassName generatedClassFQNC = ClassName.get(targetClassFQNC.packageName(), generatedClassName); // generated Class FQNC

        List<ExecutableElement> methods = elementUtil.getMethods(MetaData.class);
        List<FieldSpec> fields = createFields(methods, classMetaData, targetClassFQNC, interfaceFQNC, generatedClassFQNC);
        List<MethodSpec> methodSpecs = methods
                .stream()
                .map(s -> createMethod(s, targetClassFQNC))
                .collect(Collectors.toList());

        TypeSpec buildClass = TypeSpec.classBuilder(generatedClassFQNC)
                .addModifiers(Modifier.PUBLIC, Modifier.FINAL)
                .addSuperinterface(ParameterizedTypeName.get(interfaceFQNC, targetClassFQNC))
                .addStaticBlock(createStaticBlock(targetClassFQNC))
                .addFields(fields)
                .addMethod(createConstructor())
                .addMethods(methodSpecs)
                .build();

        return JavaFile.builder(targetClassFQNC.packageName(), buildClass).build();
    }

    private MethodSpec createConstructor() {
        return MethodSpec.constructorBuilder()
                .addModifiers(Modifier.PRIVATE)
                .build();
    }

    private List<FieldSpec> createFields(List<ExecutableElement> methods,
                                         ClassMetaData classMetaData,
                                         ClassName targetClassFQNC,
                                         ClassName interfaceFQNC,
                                         ClassName generatedClassFQNC) {
        Map<String, ExecutableElement> methodMap = methods.stream()
                .collect(Collectors.toMap(
                        method -> method.getSimpleName().toString(),
                        method -> method,
                        (o1, o2) -> o1));

        TypeMirror sheetStyleConfigurer = classMetaData.getSheetStyleConfigurerTypeMirror();
        String[] csvConfig = classMetaData.getCsvConfig();
        List<TypeMirror> headerStyles = classMetaData.getFieldList(FieldAnnotationMetaData :: getHeaderStyleConfigurerTypeMirror);
        List<TypeMirror> dataStyles = classMetaData.getFieldList(FieldAnnotationMetaData :: getDataStyleConfigurerTypeMirror);
        List<String> headerNames = classMetaData.getFieldList(FieldAnnotationMetaData :: getHeaderName);
        List<String> cellFormats = classMetaData.getFieldList(FieldAnnotationMetaData :: getCellFormat);
        List<String> defaultValues = classMetaData.getFieldList(FieldAnnotationMetaData :: getDefaultValue);
        List<String> getters = classMetaData.getFieldList(FieldAnnotationMetaData :: getGetterMethod);

        return Arrays.asList(
                createInstanceField(interfaceFQNC, generatedClassFQNC, targetClassFQNC),
                createCSVConfigField(csvConfig),
                createSheetStyleField(methodMap.get("getSheetStyle"), sheetStyleConfigurer),
                createCellStyleConfigList(methodMap.get("getHeaderStyleConfigList"), headerStyles),
                createCellStyleConfigList(methodMap.get("getDataStyleConfigList"), dataStyles),
                createStringList(methodMap.get("getColumnNameList"), headerNames),
                createStringList(methodMap.get("getDefaultValueList"), defaultValues),
                createStringList(methodMap.get("getCellFormatList"), cellFormats),
                createFunctionList(methodMap.get("getGetterFunctionList"), getters, targetClassFQNC));
    }

    private FieldSpec createInstanceField(ClassName interfaceFQNC,
                                          ClassName generatedClassFQNC,
                                          ClassName targetClassFQNC) {
        return FieldSpec.builder(ParameterizedTypeName.get(interfaceFQNC, targetClassFQNC), "INSTANCE")
                .addModifiers(Modifier.PRIVATE, Modifier.STATIC, Modifier.FINAL)
                .initializer("new $T()", generatedClassFQNC)
                .build();
    }

    private FieldSpec createSheetStyleField(ExecutableElement method, TypeMirror styleTypeMirror) {
        TypeName fieldType = elementUtil.getTypeName(method.getReturnType(), null);
        String fieldName = elementUtil.generateFieldName(method);
        CodeBlock initBlock = CodeBlock.builder()
                .add("null")
                .build();
        if (Objects.nonNull(styleTypeMirror)) {
            initBlock = CodeBlock.builder()
                    .add("new $T()", styleTypeMirror)
                    .build();
        }

        return FieldSpec.builder(fieldType, fieldName, Modifier.PRIVATE, Modifier.STATIC, Modifier.FINAL)
                .initializer(initBlock)
                .build();
    }

    private FieldSpec createFunctionList(ExecutableElement method, List<String> list, ClassName typeVar) {
        TypeName fieldType = elementUtil.getTypeName(method.getReturnType(), typeVar);
        String fieldName = elementUtil.generateFieldName(method);

        return FieldSpec.builder(fieldType, fieldName, Modifier.PRIVATE, Modifier.STATIC, Modifier.FINAL)
                .initializer(initializerFunctionList(list, typeVar))
                .build();
    }

    private CodeBlock initializerFunctionList(List<String> list, ClassName typeName) {
        CodeBlock.Builder builder = CodeBlock.builder()
                .add("$T.unmodifiableList($T.asList(", Collections.class, Arrays.class);
        int size = list.size();
        for (int i = 0; i < size - 1; i++) {
            builder.add("$T :: $L", typeName, list.get(i));
            builder.add(", ");
        }
        builder.add("$T :: $L", typeName, list.get(size - 1));
        builder.add("))");
        return builder.build();
    }

    private FieldSpec createCellStyleConfigList(ExecutableElement method, List<TypeMirror> list) {
        TypeName fieldType = elementUtil.getTypeName(method.getReturnType(), null);
        String fieldName = elementUtil.generateFieldName(method);
        return FieldSpec.builder(fieldType, fieldName, Modifier.PRIVATE, Modifier.STATIC, Modifier.FINAL)
                .initializer(initializerClassTypeList(list))
                .build();
    }

    private CodeBlock initializerClassTypeList(List<TypeMirror> list) {
        CodeBlock.Builder builder = CodeBlock.builder()
                .add("$T.unmodifiableList($T.asList(", Collections.class, Arrays.class);
        int size = list.size();
        for (int i = 0; i < size - 1; i++) {
            TypeMirror val = list.get(i);
            if (val == null) {
                builder.add("null");
            } else {
                builder.add("new $T()", val);
            }
            builder.add(", ");
        }
        TypeMirror val = list.get(size - 1);
        if (val == null) {
            builder.add("null");
        } else {
            builder.add("new $T()", val);
        }
        builder.add("))");
        return builder.build();
    }

    private FieldSpec createStringList(ExecutableElement method, List<String> list) {
        TypeName fieldType = elementUtil.getTypeName(method.getReturnType(), null);
        String fieldName = elementUtil.generateFieldName(method);
        return FieldSpec.builder(fieldType, fieldName, Modifier.PRIVATE, Modifier.STATIC, Modifier.FINAL)
                .initializer(initializerStringList(list))
                .build();
    }

    private CodeBlock initializerStringList(List<String> list) {
        CodeBlock.Builder builder = CodeBlock.builder()
                .add("$T.unmodifiableList($T.asList(", Collections.class, Arrays.class);
        int size = list.size();
        for (int i = 0; i < size - 1; i++) {
            String val = list.get(i);
            if (val == null) {
                builder.add("null");
            } else {
                builder.add("$S", val);
            }
            builder.add(", ");
        }
        String val = list.get(size - 1);
        if (val == null) {
            builder.add("null");
        } else {
            builder.add("$S", val);
        }
        builder.add("))");
        return builder.build();
    }
    private CodeBlock createStaticBlock(ClassName targetClassFQNC) {
        return CodeBlock.builder()
                .addStatement("$T.register($T.class, INSTANCE)", RegisterManager.class, targetClassFQNC)
                .build();
    }
    private FieldSpec createCSVConfigField(String[] config) {
        return FieldSpec.builder(String[].class, "csvConfig", Modifier.PRIVATE, Modifier.STATIC, Modifier.FINAL)
                .initializer(CodeBlock.of("new $T{$S, $S, $S}", String[].class, config[0], config[1], config[2]))
                .build();
    }
    private MethodSpec createMethod(ExecutableElement method, TypeName typeVar) {
        String methodName = method.getSimpleName().toString();
        TypeName returnType = elementUtil.getTypeName(method.getReturnType(), typeVar);
        String fieldName = elementUtil.generateFieldName(method);
        return createGetterMethod(methodName, fieldName, returnType);
    }

    private MethodSpec createGetterMethod(String methodName, String fieldName, TypeName returnType) {
        return MethodSpec.methodBuilder(methodName)
                .addAnnotation(Override.class)
                .addModifiers(Modifier.PUBLIC)
                .returns(returnType)
                .addStatement("return $N", fieldName)
                .build();
    }
}