package com.app.model;

import java.util.Date;



public class Employee {
	private String name;
	private String email;
	private Date dateOfBirth;
	private double salary;

	public Employee() {
		// TODO Auto-generated constructor stub
	}

	public Employee(String name, String email, Date dateOfBirth, double salary) {
		this.name = name;
		this.email = email;
		this.dateOfBirth = dateOfBirth;
		this.salary = salary;
	}

	public String getName() {
		return name;
	}

	public void setName(String name) {
		this.name = name;
	}

	public String getEmail() {
		return email;
	}

	public void setEmail(String email) {
		this.email = email;
	}

	public Date getDateOfBirth() {
		return dateOfBirth;
	}

	public void setDateOfBirth(Date dateOfBirth) {
		this.dateOfBirth = dateOfBirth;
	}

	public double getSalary() {
		return salary;
	}

	public void setSalary(double salary) {
		this.salary = salary;
	}

	@Override
	public String toString() {
		return "Employee [name=" + name + ", email=" + email + ", dateOfBirth=" + dateOfBirth + ", salary=" + salary
				+ "]";
	}

}
