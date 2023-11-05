# celper
#### *Celper는 java에서 POI를 다룰 때 불편함을 해소하고자 합니다.*
#### *Celper는 java에서 Excel을 다룰 때 "@" Annotation을 이용한 방법을 제공합니다.*
![기본개념](https://user-images.githubusercontent.com/66084125/280513712-db209cb0-0448-4ae3-a72a-2181f64ad3ef.png)

## 목차
1. [**간단하게 사용해보기**](#-간단하게-사용해보기)
2. [**더 많이 사용해보기 tutorial**](#-더-많이-사용해보기)
3. [**지원 어노테이션**](#-지원-어노테이션)


## 간단하게 사용해보기
### 목표 스프레드시트
| 이름  | 주소    | 나이  | 생년월일 |
|-----|-------|-----|------|
| 홍길동 | 00001 | 20  |37928.5|
| 김철수 | 00002 | 23  |36615.5|

### model
```java
public class StudentModel {
    @Column("이름")
    private String name;

    @Column("주소")
    private String address;

    @Column("나이")
    private int age;

    @Column("생년월일")
    private LocalDate date;

    public StudentModel() {}
    public StudentModel(String name, String address, int age, LocalDate date) {
        this.name = name;
        this.address = address;
        this.age = age;
        this.date = date;
    }
    // getter,setters ...
}
```
### Excel Service
```java
public class ExcelService {

    List<StudentModel> findByStudentModels(){
        List<StudentModel> studentModels = new ArrayList<>();

        LocalDate date1 = LocalDate.of(2003, 11, 03);
        LocalDate date2 = LocalDate.of(2000, 3, 30);

        StudentModel student1 = new StudentModel("홍길동", "00001", LocalDate.now().getYear() - date1.getYear(), date1);
        StudentModel student2 = new StudentModel("김철수", "00002", LocalDate.now().getYear() - date2.getYear(), date2);

        studentModels.add(student1);
        studentModels.add(student2);
        return studentModels;
    }

    
    void modelToSheetService() throws IOException {
        ExcelWorkBook excelWorkBook = new ExcelWorkBook(WorkBookType.XSSF);

        List<StudentModel> byStudentModels = findByStudentModels();
        excelWorkBook.createSheet().modelToSheet(byStudentModels);
        
        OutputStream outputStream = TestSupport.workBookOutput("Simple-Mapping-Student.xlsx");
        excelWorkBook.write(outputStream);
    }
}
```

## 더 많이 사용해보기
- [**go to tutorials**](https://github.com/jeongjinuk/celper/tree/main/src/test/java/org/celper/tutorial)
  - [간단하게 사용하고 싶어요! - simple_mapping](https://github.com/jeongjinuk/celper/tree/main/src/test/java/org/celper/tutorial/simple_mapping_tutorial)
  - [사용자 지정형식은 어떻게 지정하지? - cell_format](https://github.com/jeongjinuk/celper/tree/main/src/test/java/org/celper/tutorial/cellformat_tutorial)
  - [null일때 지정값을 넣고 싶은데? - default_value](https://github.com/jeongjinuk/celper/tree/main/src/test/java/org/celper/tutorial/default_value_tutorial)
  - [컬럼명을 제외하고 싶은데? - excluded_header_option](https://github.com/jeongjinuk/celper/tree/main/src/test/java/org/celper/tutorial/exculded_header_option_tutorial)
  - [스프레드시트에서 DTO로 매핑하고 싶은데.. 컬럼 이름이 다른데 어떻게하지? - import_name_option](https://github.com/jeongjinuk/celper/tree/main/src/test/java/org/celper/tutorial/import_name_option_tutorial)
  - [여러 DTO 리스트를 하나의 시트로 매핑하고 싶은데 어떻게 하지? - multimodel_mapping](https://github.com/jeongjinuk/celper/tree/main/src/test/java/org/celper/tutorial/multimodel_mapping_tutorial)
  - [컬럼의 삽입순서를 바꾸고 싶은데 어떻게 할까? - priority_option](https://github.com/jeongjinuk/celper/tree/main/src/test/java/org/celper/tutorial/priority_option_tutorial)
  - [엑셀의 스타일을 지정하고 싶은데 어떻게 하지? - style](https://github.com/jeongjinuk/celper/tree/main/src/test/java/org/celper/tutorial/style_tutorial)

---
## 지원 어노테이션
<details>
<summary> @Column </summary> 

``` java
    // 제공 옵션 및 설명
    // importNameOptions은 가장 먼저 찾게된 컬럼명을 기준으로 매핑작업을 진행합니다.
    // 그로인하여 하나의 스프레드시트에 여러개의 동일 컬럼이 존재할 경우 원하지 않는 결과가 나올 수 있습니다.
    
    // 기본 컬럼명을 정의할 수 있습니다.
    String value(); 
    
    // 스프레드시트에서 DTO 모델 매핑 시 컬럼명이 달라질 경우 사용할 수 있습니다.
    String[] importNameOptions() default {""};
    
     // 스프레드시트에 컬럼 삽입 우선순위를 지정할 수 있습니다. 높을 수록 가장 먼저 삽입됩니다. 만약 우선순위가 같다면 컬럼명을 기준으로 우선순위를 정합니다. 
    int priority() default 0;
```
</details>
<details>
<summary> @CellFormat </summary>

``` java
    // 제공 옵션 및 설명
    // 만약 customFormat과 builtinFormat을 동시에 지정할 경우 customFormat이 적용됩니다.
     
    // 기본적으로 제공하는 format을 이용할 수 있습니다.
    BuiltinCellFormatType builtinFormat() default BuiltinCellFormatType.GENERAL;
    
    // 커스텀한 format을 제작할 수 있습니다.
    String customFormat() default "";
```
</details>
<details>
<summary> @DefaultValue </summary>

``` java
    // 제공 옵션 및 설명
    // 만약 숫자일 경우 삽입되지 않습니다.
    // 이유는 null은 존재하지 않는 객체일 경우로 판단이 가능하지만, 숫자의 경우 double로 치환하는 과정에서 0이 됩니다.
     
    // null이 발생할 경우 대치되는 문자를 미리 지정할 수 있습니다.
    String value();
```
</details>
<details>
<summary> @SheetStyle </summary>

- [**tutorial style 부분을 참고하세요**](#-더-많이-사용해보기)

| 이름  | 주소    | 나이  | 생년월일 |
|-----|-------|-----|------|
| 홍길동 | 00001 | 20  |37928.5|
| 김철수 | 00002 | 23  |36615.5|

``` java
    // 제공 옵션
    // 스프레드시트에 대한 스타일을 설정해야할 경우 사용할 수 있습니다.
    // 스프레드시트에 대한 스타일을 설정해야할 경우 아래의 SheetStyleConfigurer를 구현하면 됩니다.
    // 헤더, 데이터 영역과 관계없이 모두 적용됩니다.
    Class<? extends SheetStyleConfigurer> value();
    
    ...
    public interface SheetStyleConfigurer extends StyleConfigurer<SheetStyleBuilder> {}
        
```
</details>
<details>
<summary> @ColumnStyle </summary>

- [**tutorial style 부분을 참고하세요**](#-더-많이-사용해보기)

| 이름  | 주소    | 나이  | 생년월일 |
|-----|-------|-----|------|
| 홍길동 | 00001 | 20  |37928.5|
| 김철수 | 00002 | 23  |36615.5|

``` java
    // 제공 옵션
    
    // 이름, 주소, 나이, 생년월일에 해당되는 부분의 스타일을 정의할 수 있는 옵션입니다. 
    Class<? extends CellStyleConfigurer> headerAreaStyle() default _NoCellStyle.class;
    
    // 실제 컬럼명에 해당하는 데이터에 스타일을 정의할 수 있는 옵션입니다.
    Class<? extends CellStyleConfigurer> dataAreaStyle() default _NoCellStyle.class;
    
    ...
    public interface CellStyleConfigurer extends StyleConfigurer<CellStyleBuilder> {}
```
</details>



