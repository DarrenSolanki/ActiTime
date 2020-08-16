package com.actitime.tests;

import com.actitime.generic.ExcelData;

public class ProfileTest {

	public static void main(String[] args) {
		// TODO Auto-generated method stub

		String username = ExcelData.getData("./data/input.xlsx", "TC01", 1, 0);
		String password = ExcelData.getData("./data/input.xlsx", "TC01", 1, 1);
		String loginTitle = ExcelData.getData("./data/input.xlsx", "TC01", 1, 2);
		String enterTimeTrackTitle = ExcelData.getData("./data/input.xlsx", "TC01", 1, 3);
		
		System.out.println(username);
		System.out.println(password);
		System.out.println(loginTitle);
		System.out.println(enterTimeTrackTitle);

		
	}

}
