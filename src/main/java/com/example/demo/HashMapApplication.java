package com.example.demo;

import java.util.HashMap;

import org.springframework.boot.SpringApplication;
import org.springframework.boot.autoconfigure.SpringBootApplication;

@SpringBootApplication
public class HashMapApplication {

	public static void main(String[] args) {
		
		String excelName="Test";
		
		HashMap<String,String> ColoumnHeader1= new HashMap<String, String>();
		ColoumnHeader1.put("Name", "Name");
		ColoumnHeader1.put("Width", "16000");
		ColoumnHeader1.put("FontName", "Arial");
		ColoumnHeader1.put("FontHeight", "16");
		ColoumnHeader1.put("Bold", "true");
		
		HashMap<String,String> ColoumnHeader2= new HashMap<String, String>();
		ColoumnHeader2.put("Name", "Address");
		ColoumnHeader2.put("Width", "18000");
		ColoumnHeader2.put("FontName", "Arial");
		ColoumnHeader2.put("FontHeight", "16");
		ColoumnHeader2.put("Bold", "true");
		
		HashMap<String,String> ColoumnHeader3= new HashMap<String, String>();
		ColoumnHeader3.put("Name", "Age");
		ColoumnHeader3.put("Width", "6000");
		ColoumnHeader3.put("FontName", "Arial");
		ColoumnHeader3.put("FontHeight", "16");
		ColoumnHeader3.put("Bold", "false");
		
		
		HashMap<String,HashMap<String,String>> ColumnHeaders=new HashMap<>();
		ColumnHeaders.put("1",ColoumnHeader1);
		ColumnHeaders.put("2",ColoumnHeader2);
		ColumnHeaders.put("3",ColoumnHeader3);

		
//		Object [] data= {"thuva","manipay",22,"thivi","uduvil",27,"kavi","chunnakam",26,"ravi","uk",54};
		String [] data= {"thuva","manipay","23","thivi","uduvil","27","kavi","chunnakam","26","ravi","uk","54"};
		HashMapDemo1.writeExcel(excelName, ColumnHeaders,data);
		SpringApplication.run(HashMapApplication.class, args);
	}

}
