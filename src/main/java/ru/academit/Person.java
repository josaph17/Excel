package ru.academit;

public class Person {
    private int age;
    private String familyName;
    private String name;
    private String phoneNumber;

    public int getAge() {
        return age;
    }

    public void setAge(int age) {
        this.age = age;
    }

    public String getFamilyName() {
        return familyName;
    }

    public void setFamilyName(String familyName) {
        this.familyName = familyName;
    }

    public String getName() {
        return name;
    }

    public void setName(String name) {
        this.name = name;
    }

    public String getPhoneNumber() {
        return phoneNumber;
    }

    public void setPhoneNumber(String phoneNumber) {
        this.phoneNumber = phoneNumber;
    }

    public Person(int age, String familyName, String name, String phoneNumber) {
        this.age = age;
        this.familyName = familyName;
        this.name = name;
        this.phoneNumber = phoneNumber;
    }

    @Override
    public String toString() {
        return "{" + familyName + ", " + name + ", " + age + ", " + phoneNumber + '}';
    }
}
