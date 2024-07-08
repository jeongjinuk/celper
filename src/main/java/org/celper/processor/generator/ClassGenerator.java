//package org.celper.processor.generator;
//
//import com.squareup.javapoet.*;
//import org.celper.MetaData;
//import org.celper.core.style.SheetStyleConfigurer;
//import org.celper.processor.meta.ClassMetaData;
//import org.celper.processor.util.ElementUtil;
//
//import javax.lang.model.element.Modifier;
//import javax.lang.model.element.Name;
//import javax.lang.model.element.VariableElement;
//
///**
// * javapoet
// * $T = Type
// * $S = 문자열
// * $L = 리터럴 숫자, 문자 등등
// * $N = 변수명으로 접근가능
// *
// * 필요
// * Element to TypeElement convert
// * getQualifiedName -> FQCN
// * getSimepleName -> ClassName
// *
// * FQCN 풀패키지 + class이름
// * className = class 이름
// * generateClassFQCN = (FQCN - className) + "__Generate__Model<FQCN>";
// * generatedPackage = (FQCN - className) + generateClassFQCN;
// *
// * static 블록 생성기
// *      여기서 RegisterManager.put(Class.class, new Generated___Model()); 이런식
// *
// *
// * 생성자 생성기 constructor
// *     여기서 필드 초기화
// *
// * statement 생성기
// *     this.sheetStyle = builder -> {};
// *     this.headerStyleList = Collections.unmodifiableList(Arrays.asList());
// *     this.dataStyleList = Collections.unmodifiableList(Arrays.asList());
// *     this.columnNameList = Collections.unmodifiableList(Arrays.asList());
// *     this.defaultValueList = Collections.unmodifiableList(Arrays.asList());
// *     this.getterList = Collections.unmodifiableList(Arrays.asList(DTO::getAge, ...));
// *
// * 필드 생성기
// *     private SheetStyleConfigurer sheetStyle;
// *     private List<CellStyleConfigurer> headerStyleList;
// *     private List<CellStyleConfigurer> dataStyleList;
// *     private List<String> columnNameList;
// *     private List<Object> defaultValueList;
// *     private List<Function<DTO, Object>> getterList;
// *
// *
// * 메서드 생성기
// *     public SheetStyleConfigurer getSheetStyle();
// *     public List<CellStyleConfigurer> getHeaderStyle();
// *     public List<CellStyleConfigurer> getDataStyle();
// *     public List<String> getColumnNames();
// *     public List<Object> getDefaultValues();
// *     public List<Function<T, Object>> getGetterList();
// */
//public class ClassGenerator implements Generator<ClassMetaData, TypeSpec>{
//
//    /**
//     * 생성 클래스 이름 : [DTO Class Name]_Meta_Data
//     * 생성 위치 : [DTO Class package].[DTO Class Name]_Meta_Data;
//     * 필요한거 - 타겟 인터페이스의 메서드들
//     *         - 타겟 인터페이스에 만들어질 필드들
//     *         - 생성자
//     *         - 이름
//     * 조건? 무조건 getXXXX()
//     * 1. getter
//     * 2. collection.unmod
//     * 3. List
//     * 4. no param
//     *
//     */
//
//
//    @Override
//    public TypeSpec generate(ClassMetaData classMetaData) {
//        Name qualifiedName = classMetaData.getClazz().getQualifiedName(); //fqcn
//        ParameterizedTypeName.get()
//        return null;
//    }
//
//    private MethodSpec getConstructor(){
//        return MethodSpec.constructorBuilder()
//                .addModifiers(Modifier.PRIVATE)
//                .build();
//    }
//
//
//    private MethodSpec getMetaModelOverrideMethod(String methodName, TypeName typeName){
//        return MethodSpec.methodBuilder(methodName)
//                .addAnnotation(Override.class)
//                .addModifiers(Modifier.PUBLIC)
//                // TODO List<?> 일수 있고, 일반 객체일 수 있음
//                .returns(SheetStyleConfigurer.class)
//                .
//    }
//
//
//
//}
