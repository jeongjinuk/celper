package org.celper.tutorial.import_name_option_tutorial;

import org.celper.annotations.Column;

import java.time.LocalDate;
import java.util.Objects;

public class StudentModel {
    @Column(value = "이름", importNameOptions = {"성명"})
    private String name;

    @Column("주소")
    private String address;

    @Column("나이")
    private int age;

    @Column("생년월일")
    private LocalDate date;

    public StudentModel() {
    }

    public StudentModel(String name, String address, int age, LocalDate date) {
        this.name = name;
        this.address = address;
        this.age = age;
        this.date = date;
    }

    public String getName() {
        return name;
    }

    public String getAddress() {
        return address;
    }

    public int getAge() {
        return age;
    }

    public LocalDate getDate() {
        return date;
    }

    public void setName(String name) {
        this.name = name;
    }

    public void setAddress(String address) {
        this.address = address;
    }

    public void setAge(int age) {
        this.age = age;
    }

    public void setDate(LocalDate date) {
        this.date = date;
    }

    @Override
    public String toString() {
        return "StudentModel{" +
                "name='" + name + '\'' +
                ", address='" + address + '\'' +
                ", age=" + age +
                ", date=" + date +
                '}';
    }

    @Override
    public boolean equals(Object o) {
        if (this == o) return true;
        if (o == null || getClass() != o.getClass()) return false;
        StudentModel that = (StudentModel) o;
        return age == that.age && name.equals(that.name) && address.equals(that.address) && date.equals(that.date);
    }

    @Override
    public int hashCode() {
        return Objects.hash(name, address, age, date);
    }
}
