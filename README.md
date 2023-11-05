# celper
#### *Celper는 java에서 POI를 다룰 때 불편함을 해소하고자 합니다.*
#### *Celper는 java에서 Excel을 다룰 때 "@" Annotation을 이용한 방법을 제공합니다.*
![기본개념](https://user-images.githubusercontent.com/66084125/280513712-db209cb0-0448-4ae3-a72a-2181f64ad3ef.png)

## 간단하게 사용해보기

---
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
## 기능
| 이름  | 주소    | 나이  | 생년월일 |
|-----|-------|-----|------|
| 홍길동 | 00001 | 20  |37928.5|
| 김철수 | 00002 | 23  |36615.5|
