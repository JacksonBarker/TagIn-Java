package tagIn;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.util.Scanner;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class tagIn {
	
	public static String inpt;
	public static String wbLoc = "W:\\tagIn\\demosheet.xlsx";
	
	public static void main(String[] args) {
		Scanner sc = new Scanner(System.in);
		inpt = sc.nextLine();
		if (inpt.contains("newuser")) {
			System.out.println("What is the users name?");
			inpt = sc.nextLine();
			if (inpt.contains("cancel") != true) {
				String newname = inpt;
				System.out.println("What is the users UUID?");
				inpt = sc.nextLine();
				if (inpt.contains("cancel") != true) {
					String newuuid = inpt;
					CreateUser(newname, newuuid);
				}
			}
		}
	}

	public static void CreateUser(String cName, String cUUID) {
		Workbook wb = null;
		int cr = 0;
		try {
			FileInputStream fis = new FileInputStream(wbLoc);
			wb = new XSSFWorkbook(fis);
		} catch (FileNotFoundException e2) {
			e2.printStackTrace();
		} catch (IOException e3) {
			e3.printStackTrace();
		}
		Sheet sheet = wb.getSheetAt(0);
		boolean created = false;
		while (created = false) {
			try {
				Row row = sheet.getRow(cr);
				Cell cell = row.getCell(0);
				cr++;
				System.out.println(cr);
			} catch (NullPointerException e4) {
				System.out.println("hello");
				Row row = sheet.getRow(cr);
				Cell cell = row.createCell(0);
			    cell.setCellValue(cName);
			    cell = row.createCell(1);
			    cell.setCellValue(cUUID);
			    cell = row.createCell(2);
			    cell.setCellValue("out");
			    created = true;
			}
		}
	}

}