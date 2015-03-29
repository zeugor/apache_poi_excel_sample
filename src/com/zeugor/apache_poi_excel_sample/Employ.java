package com.zeugor.apache_poi_excel_sample;

public class Employ {
	private int id;
	private String name;
	private int salary;

	public int getId() {
		return id;
	}

	public String getName() {
		return name;
	}

	public int getSalary() {
		return salary;
	}

	public void setId(int id) {
		this.id = id;
	}

	public void setName(String name) {
		this.name = name;
	}

	public void setSalary(int salary) {
		this.salary = salary;
	}

	@Override
	public String toString() {
		return ">> Employ [id: " + id + ", name: " + name + ", salary: " + salary + "]";
	}
}
