package org.celper.processor.generator;

import com.squareup.javapoet.*;
import org.celper.MetaData;
import org.celper.core.style.CellStyleConfigurer;
import org.celper.core.style.SheetStyleConfigurer;
import org.celper.processor.meta.ClassMetaData;
import org.celper.processor.meta.FieldAnnotationMetaData;
import org.celper.processor.util.ElementUtil;
import org.celper.register.RegisterManager;

import javax.lang.model.element.ExecutableElement;
import javax.lang.model.element.Modifier;
import java.util.*;
import java.util.stream.Collectors;

/**
 * javapoet
 * $T = Type
 * $S = 문자열
 * $L = 리터럴 숫자, 문자 등등
 * $N = 변수명으로 접근가능
 * <p>
 * 필요
 * Element to TypeElement convert
 * getQualifiedName -> FQCN
 * getSimepleName -> ClassName
 * <p>
 * FQCN 풀패키지 + class이름
 * className = class 이름
 * generateClassFQCN = (FQCN - className) + "__Generate__Model<FQCN>";
 * generatedPackage = (FQCN - className) + generateClassFQCN;
 * <p>
 * static 블록 생성기
 * 여기서 RegisterManager.put(Class.class, new Generated___Model()); 이런식
 * <p>
 * <p>
 * 생성자 생성기 constructor
 * 여기서 필드 초기화
 * <p>
 * statement 생성기
 * this.sheetStyle = builder -> {};
 * this.headerStyleList = Collections.unmodifiableList(Arrays.asList());
 * this.dataStyleList = Collections.unmodifiableList(Arrays.asList());
 * this.columnNameList = Collections.unmodifiableList(Arrays.asList());
 * this.defaultValueList = Collections.unmodifiableList(Arrays.asList());
 * this.getterList = Collections.unmodifiableList(Arrays.asList(DTO::getAge, ...));
 * <p>
 * 필드 생성기
 * private SheetStyleConfigurer sheetStyle;
 * private List<CellStyleConfigurer> headerStyleList;
 * private List<CellStyleConfigurer> dataStyleList;
 * private List<String> columnNameList;
 * private List<Object> defaultValueList;
 * private List<Function<DTO, Object>> getterList;
 * <p>
 * <p>
 * 메서드 생성기
 * public SheetStyleConfigurer getSheetStyle();
 * public List<CellStyleConfigurer> getHeaderStyle();
 * public List<CellStyleConfigurer> getDataStyle();
 * public List<String> getColumnNames();
 * public List<Object> getDefaultValues();
 * public List<Function<T, Object>> getGetterList();
 */
public class ClassGenerator implements Generator<ClassMetaData, JavaFile> {

    private static final String suffix = "_MetaData";

    private final ElementUtil elementUtil;

    public ClassGenerator(ElementUtil elementUtil) {
        this.elementUtil = elementUtil;
    }

    @Override
    public JavaFile generate(ClassMetaData classMetaData) {

        /**
         * 필요한거
         *
         * 1. 풀패키지 경로 -> dto 클래스 경로에서 dto만 제외
         * 2. dto 클래스 이름
         * 3. 만들어질 generate 클래스
         * 4. generate_MetaData -> 타입 생성 ParameterizedTypeName.get(1 경로, 2dto + suffix)
         * 5. TypeSpec에서 supperInterface(4)
         *
         * private static final org.example.Temp<org.example.DTO> INSTANCE = new org.example.DTO_Generate(); 이렇게 나오는 것 까진 확인
         */
        ClassName targetClassFQNC = ClassName.get(classMetaData.getClazz());
        ClassName interfaceFQNC = ClassName.get(MetaData.class);
        String generatedClassName = targetClassFQNC.simpleName() + suffix;
        ClassName generatedClassFQNC = ClassName.get(targetClassFQNC.packageName(), generatedClassName);


        /**
         *
         * 필드 타입 == 메서드 리턴
         * 재사용
         * Map<MethodSpec, fieldName> 형으로 반환
         * -> 요렇게 해서 stream으로 파싱? 이게 맞네
         *
         * 들고 오는거 까지는 ok
         * ClassMetaData의 필드랑 매핑 시키는건? 걍 박으면 의미가 하나도 없는데
         * 구현까진 만들 수 있는데 의미가 없는데 반자동화 수준?
         * 맞음?
         * 뒤에는 바뀔 수 있다? 자동화시킨다고 해도 명칭이 달라서 불가능 어케 만들건데
         * ClassMetaData
         * $N
         *
         */
        List<ExecutableElement> methods = elementUtil.getMethods(MetaData.class);
        List<FieldSpec> fields = createFields(methods, classMetaData, targetClassFQNC, interfaceFQNC, generatedClassFQNC);
        List<MethodSpec> methodSpecs = methods
                .stream()
                .map(s -> createGetterMethod(s, targetClassFQNC))
                .collect(Collectors.toList());

        /**
         *
         *        TypeSpec generatedMetaDataClass = TypeSpec.classBuilder(generatedMetaData)
         *                 .addModifiers(Modifier.PUBLIC, Modifier.FINAL)
         *                 .addSuperinterface(ParameterizedTypeName.get(metaData, dto)) // 상속받게 해야함
         *                 .addField(instanceField)
         *                 .build();
         */
        TypeSpec buildClass = TypeSpec.classBuilder(generatedClassFQNC)
                .addModifiers(Modifier.PUBLIC, Modifier.FINAL)
                .addSuperinterface(ParameterizedTypeName.get(interfaceFQNC, targetClassFQNC))
                .addStaticBlock(createStaticBlock(targetClassFQNC))
                .addFields(fields)
                .addMethod(createConstructor())
                .addMethods(methodSpecs)
                .build();
        System.out.println();
        System.out.println(targetClassFQNC.packageName() + buildClass + "--------");
        System.out.println();
        JavaFile build = JavaFile.builder(targetClassFQNC.packageName(), buildClass).build();
        return JavaFile.builder(targetClassFQNC.packageName(), buildClass).build();
    }

    private MethodSpec createConstructor() {
        return MethodSpec.constructorBuilder()
                .addModifiers(Modifier.PRIVATE)
                .build();
    }

    private List<FieldSpec> createFields(List<ExecutableElement> methods,
                                         ClassMetaData classMetaData,
                                         ClassName targetClass,
                                         ClassName interfaceFQNC,
                                         ClassName generatedClassFQNC) {
        Map<String, ExecutableElement> methodMap = methods.stream()
                .collect(Collectors.toMap(
                        method -> method.getSimpleName().toString(),
                        method -> method,
                        (o1, o2) -> o1));

        Class<SheetStyleConfigurer> sheetStyleConfigurer = classMetaData.getSheetStyleConfigurer();
        List<Class<CellStyleConfigurer>> headerStyles = classMetaData.getFieldList(FieldAnnotationMetaData :: getHeaderStyleConfigurer);
        List<Class<CellStyleConfigurer>> dataStyles = classMetaData.getFieldList(FieldAnnotationMetaData :: getDataStyleConfigurer);
        List<String> headerNames = classMetaData.getFieldList(FieldAnnotationMetaData :: getHeaderName);
        List<String> cellFormats = classMetaData.getFieldList(FieldAnnotationMetaData :: getCellFormat);
        List<String> getters = classMetaData.getFieldList(FieldAnnotationMetaData :: getGetterMethod);
        List<String> defaultValues = classMetaData.getFieldList(FieldAnnotationMetaData :: getDefaultValue);

        return Arrays.asList(
                createInstanceField(interfaceFQNC, generatedClassFQNC),
                createSheetStyleField(methodMap.get("getSheetStyle"), sheetStyleConfigurer),
                createCellStyleConfigList(methodMap.get("getHeaderStyleConfigList"), headerStyles),
                createCellStyleConfigList(methodMap.get("getDataStyleConfigList"), dataStyles),
                createStringList(methodMap.get("getColumnNameList"), headerNames),
                createStringList(methodMap.get("getDefaultValueList"), defaultValues),
                createStringList(methodMap.get("getCellFormatList"), cellFormats),
                createFunctionList(methodMap.get("getGetterFunctionList"), getters, targetClass));
    }

    private FieldSpec createInstanceField(ClassName interfaceFQNC, ClassName generatedClassFQNC) {
        return FieldSpec.builder(interfaceFQNC, "INSTANCE", Modifier.PRIVATE, Modifier.STATIC, Modifier.FINAL)
                .initializer("new $T()", generatedClassFQNC)
                .build();
    }

    private FieldSpec createSheetStyleField(ExecutableElement method, Class<?> clazz) {
        TypeName fieldType = elementUtil.getTypeName(method.getReturnType(), null);
        String fieldName = elementUtil.generateFieldName(method);
        CodeBlock initBlock = CodeBlock.builder()
                .add("null")
                .build();

        if (Objects.nonNull(clazz)) {
            initBlock = CodeBlock.builder()
                    .add("new $T()", ClassName.get(clazz))
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

    private FieldSpec createCellStyleConfigList(ExecutableElement method, List<Class<CellStyleConfigurer>> list) {
        TypeName fieldType = elementUtil.getTypeName(method.getReturnType(), null);
        String fieldName = elementUtil.generateFieldName(method);
        return FieldSpec.builder(fieldType, fieldName, Modifier.PRIVATE, Modifier.STATIC, Modifier.FINAL)
                .initializer(initializerClassTypeList(list))
                .build();
    }

    private CodeBlock initializerClassTypeList(List<Class<CellStyleConfigurer>> list) {
        CodeBlock.Builder builder = CodeBlock.builder()
                .add("$T.unmodifiableList($T.asList(", Collections.class, Arrays.class);
        int size = list.size();
        for (int i = 0; i < size - 1; i++) {
            Class<CellStyleConfigurer> val = list.get(i);
            if (val == null) {
                builder.add("null");
            } else {
                builder.add("new $T()", ClassName.get(val));
            }
            builder.add(", ");
        }
        Class<CellStyleConfigurer> val = list.get(size - 1);
        if (val == null) {
            builder.add("null");
        } else {
            builder.add("new $T()", ClassName.get(val));
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
                .addStatement("$T.put($T.class, INSTANCE)", RegisterManager.class, targetClassFQNC)
                .build();
    }


    private MethodSpec createGetterMethod(ExecutableElement method, TypeName typeVar) {
        String methodName = method.getSimpleName().toString();
        String fieldName = elementUtil.generateFieldName(method);
        TypeName returnType = elementUtil.getTypeName(method.getReturnType(), typeVar);

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
