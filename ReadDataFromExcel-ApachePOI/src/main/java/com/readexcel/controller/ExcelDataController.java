package com.readexcel.controller;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.io.InputStream;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.nio.file.StandardCopyOption;
import java.text.ParseException;
import java.text.SimpleDateFormat;
import java.util.*;

import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.core.io.ClassPathResource;
import org.springframework.stereotype.Controller;
import org.springframework.ui.Model;
import org.springframework.web.bind.annotation.GetMapping;
import org.springframework.web.bind.annotation.PostMapping;
import org.springframework.web.bind.annotation.RequestParam;
import org.springframework.web.multipart.MultipartFile;

@Controller
public class ExcelDataController {

	@GetMapping("/form")
	public String getForm() {

		return "index";
	}

	@PostMapping("/upload")
	public String uploadFile(@RequestParam("excel-file") MultipartFile file, Model m) {

		String filename = file.getOriginalFilename();
		m.addAttribute("filename", filename);
		try {

			// file path where to upload
			File filepath = new ClassPathResource("/templates").getFile();
			System.out.println(filepath);
			// get complete path using java.nio package
			Path path = Paths.get(filepath.getAbsolutePath() + File.separator + filename);
			// uplaod file
			Files.copy(file.getInputStream(), path, StandardCopyOption.REPLACE_EXISTING);

		} catch (Exception e) {
			e.printStackTrace();
		}

		return "index";
	}

	// get data
	@GetMapping("/getdata")
	public String getdata(Model m) {

		try {
			FileInputStream fis = new FileInputStream(
					new ClassPathResource("/templates/Assignment_Timecard.xlsx").getFile());
			XSSFWorkbook wb = new XSSFWorkbook(fis);
			XSSFSheet sheet = wb.getSheet("Sheet1");

			int lastRowNum = sheet.getLastRowNum();
			int lastCellNum = sheet.getRow(0).getLastCellNum();
			List<String> list=new ArrayList();
			List<String> list2=new ArrayList();
			List<String> tlist=new ArrayList();
			List<String> tlist2=new ArrayList();
			for (int r = 1; r <=lastRowNum; r++) {
				XSSFRow row = sheet.getRow(r);
				XSSFCell cell = row.getCell(5);
				XSSFCell cell1 = row.getCell(6);
				XSSFCell tcell = row.getCell(2);
				XSSFCell tOcell1 = row.getCell(3);
				SimpleDateFormat sdf=new SimpleDateFormat("MM-DD-yyyy HH:mm:ss");
				String ds1 = cell.toString();
				String ds2 = cell1.toString();
				if(ds1!=null && ds2!=null) {
					
					Date d1=new Date(ds1);
					Date d2=new Date(ds2);
					Date td1=new Date(tcell.toString());
					Date tOd2=new Date(tOcell1.toString());
					System.out.println(td1+" "+tOd2);
					long difference_In_Time
		            = d2.getTime() - d1.getTime();
					long difference_In_Timeshift
		            = tOd2.getTime() - td1.getTime();
					long difference_In_Days
		            = (difference_In_Time
		               / (1000 * 60 * 60 * 24))
		              % 365;
					long difference_In_Hours
	                = (difference_In_Timeshift
	                   / (1000 * 60 * 60))
	                  % 24;
					if(difference_In_Days>7) {
						
						
						list2.add(row.getCell(1).getStringCellValue());
						list.add(row.getCell(7).getStringCellValue());
						System.out.println(row.getCell(7).getStringCellValue());
					}
					if(difference_In_Hours<10 && difference_In_Hours>1) {
						tlist.add(row.getCell(1).getStringCellValue());
						tlist2.add(row.getCell(7).getStringCellValue());
					}
				}
			}
			m.addAttribute("pos",list);
			m.addAttribute("name",list2);
			m.addAttribute("tpos",tlist);
			m.addAttribute("tname",tlist2);
			System.out.println(fis.toString());
		} catch (IOException e) {
			e.printStackTrace();
		}
		return "index";
	}

}
