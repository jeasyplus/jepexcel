package com.jeasyplus.excel.entity;

public class TableObject {

    private String name;

    private int age;

    private String des;

    private String abc;

    public TableObject(String name, int age, String des, String abc) {
        this.name = name;
        this.age = age;
        this.des = des;
        this.abc = abc;
    }

    public String getName() {
        return name;
    }

    public void setName(String name) {
        this.name = name;
    }

    public int getAge() {
        return age;
    }

    public void setAge(int age) {
        this.age = age;
    }

    public String getDes() {
        return des;
    }

    public void setDes(String des) {
        this.des = des;
    }

    public String getAbc() {
        return abc;
    }

    public void setAbc(String abc) {
        this.abc = abc;
    }
}
