package com.ExcelReading;

import java.util.ArrayList;

import org.testng.annotations.Test;

public class TestExcel {
	
	public void callExcelFile() {
		ExcelReading ex = new ExcelReading();
		ArrayList data = ex.testData("F:\\Data files\\test.xlsx", 0);
		System.out.println(data);
	}
}
