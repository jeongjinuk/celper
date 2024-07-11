package org.celper.processor.generator;

import com.squareup.javapoet.*;
import org.celper.MetaData;
import org.celper.processor.meta.ClassMetaData;
import org.celper.processor.util.ElementUtil;
import org.celper.register.RegisterManager;

import javax.lang.model.element.ExecutableElement;
import javax.lang.model.element.Modifier;
import javax.lang.model.element.Name;
import javax.lang.model.type.TypeKind;
import java.util.List;
import java.util.Map;
import java.util.function.BiFunction;
import java.util.stream.Collectors;

/**
 * javapoet
 * $T = Type
 * $S = 문자열
 * $L = 리터럴 숫자, 문자 등등
 * $N = 변수명으로 접근가능
 *
 * 필요
 * Element to TypeElement convert
 * getQualifiedName -> FQCN
 * getSimepleName -> ClassName
 *
 * FQCN 풀패키지 + class이름
 * className = class 이름
 * generateClassFQCN = (FQCN - className) + "__Generate__Model<FQCN>";
 * generatedPackage = (FQCN - className) + generateClassFQCN;
 *
 * static 블록 생성기
 *      여기서 RegisterManager.put(Class.class, new Generated___Model()); 이런식
 *
 *
 * 생성자 생성기 constructor
 *     여기서 필드 초기화
 *
 * statement 생성기
 *     this.sheetStyle = builder -> {};
 *     this.headerStyleList = Collections.unmodifiableList(Arrays.asList());
 *     this.dataStyleList = Collections.unmodifiableList(Arrays.asList());
 *     this.columnNameList = Collections.unmodifiableList(Arrays.asList());
 *     this.defaultValueList = Collections.unmodifiableList(Arrays.asList());
 *     this.getterList = Collections.unmodifiableList(Arrays.asList(DTO::getAge, ...));
 *
 * 필드 생성기
 *     private SheetStyleConfigurer sheetStyle;
 *     private List<CellStyleConfigurer> headerStyleList;
 *     private List<CellStyleConfigurer> dataStyleList;
 *     private List<String> columnNameList;
 *     private List<Object> defaultValueList;
 *     private List<Function<DTO, Object>> getterList;
 *
 *
 * 메서드 생성기
 *     public SheetStyleConfigurer getSheetStyle();
 *     public List<CellStyleConfigurer> getHeaderStyle();
 *     public List<CellStyleConfigurer> getDataStyle();
 *     public List<String> getColumnNames();
 *     public List<Object> getDefaultValues();
 *     public List<Function<T, Object>> getGetterList();
 */
public class ClassGenerator implements Generator<ClassMetaData, TypeSpec>{

    /**
     * 생성 클래스 이름 : [DTO Class Name]_Meta_Data
     * 생성 위치 : [DTO Class package].[DTO Class Name]_Meta_Data;
     * 필요한거 - 타겟 인터페이스의 메서드들
     *         - 타겟 인터페이스에 만들어질 필드들
     *         - 생성자
     *         - 이름
     *
     *
     * 해야할거
     * TypeName 만들어주는 거.
     *  - List<Object,Map<Object,T>>> -> 이런식일때
     *  - Function<T, S> 이런식일때
     *  재귀 처리하는게 좋음
     *     private TypeName getTypeName(TypeMirror returnType){
     *         if (returnType.getKind().isPrimitive()){
     *             TypeElement typeElement = processingEnv.getTypeUtils().boxedClass((PrimitiveType) returnType);
     *             return ClassName.get(typeElement);
     *         }
     *         switch (returnType.getKind()){
     *             case DECLARED:
     *                 DeclaredType declaredType = (DeclaredType) returnType;
     *                 TypeElement typeElement = (TypeElement) declaredType.asElement();
     *                 System.out.println("TypeKind.DECLARED - " + returnType);
     *                 ClassName className = ClassName.get(typeElement);
     *
     *                 if (!declaredType.getTypeArguments().isEmpty()){
     *                     TypeName[] typeNames = declaredType.getTypeArguments()
     *                             .stream()
     *                             .map(this::getTypeName)
     *                             .toArray(TypeName[] :: new);
     *                     return ParameterizedTypeName.get(className, typeNames);
     *                 }
     *                 return className;
     *             case TYPEVAR: // T 타입
     *                 TypeVariable typeVariable = (TypeVariable) returnType;
     *                 return TypeVariableName.get(typeVariable);
     *             default:
     *                 return TypeName.get(returnType);
     *         }
     *     }
     *
     *
     *
     */

    private static final String suffix = "_MetaData";

    private final ElementUtil elementUtil;

    public ClassGenerator(ElementUtil elementUtil) {
        this.elementUtil = elementUtil;
    }

    @Override
    public TypeSpec generate(ClassMetaData classMetaData) {
        Name qualifiedName = classMetaData.getClazz().getQualifiedName(); //fqcn

        List<MethodSpec> collect = elementUtil.getMethods(MetaData.class)
                .stream()
                .map(this :: createGetterMethod)
                .collect(Collectors.toList());

        return null;
    }





    private MethodSpec createConstructor(){
        return MethodSpec.constructorBuilder()
                .addModifiers(Modifier.PRIVATE)
                .build();
    }

    private CodeBlock createStaticBlock(ClassName clazz){
        return CodeBlock.builder()
                .addStatement("$T.put($T.class, INSTANCE)", RegisterManager.class, clazz)
                .build();
    }

    private MethodSpec createGetterMethod(ExecutableElement method){
        String methodName = method.getSimpleName().toString();
        String fieldName = elementUtil.generateFieldName(method);
        TypeName returnType = null;
        return createGetterMethod(methodName,fieldName, returnType);
    }


    private MethodSpec createGetterMethod(String methodName, String fieldName, TypeName returnType){
        return MethodSpec.methodBuilder(methodName)
                .addAnnotation(Override.class)
                .addModifiers(Modifier.PUBLIC)
                .returns(returnType)
                .addStatement("return $N", fieldName)
                .build();
    }


}
