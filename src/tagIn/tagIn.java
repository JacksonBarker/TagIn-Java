package tagIn;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.Scanner;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class tagIn {
	
	public static String inpt;
	public static String wbLoc = "C:/eclipse/TagIn-Java/demosheet.xlsx";
	public static Boolean stop = false;
	
	public static void main(String[] args) {
		System.out.println("TagIn Java");
		System.out.println("Version: in development");
		System.out.println("Ready for command.");
		Scanner sc = new Scanner(System.in);
		while (!stop) {
			inpt = sc.nextLine();
			if (inpt.contains("newuser")) {
				System.out.println("Send 'cancel' at any time to cancel user creation.");
				System.out.println("What is the users name?");
				inpt = sc.nextLine();
				if (!inpt.contains("cancel")) {
					String newname = inpt;
					System.out.println("What is the users UUID?");
					inpt = sc.nextLine();
					if (!inpt.contains("cancel")) {
						String newuuid = inpt;
						System.out.println("WARNING PLEASE CHECK THAT THE ROW DOESN'T CONTAIN ANOTHER USER AND IS DIRECTLY BELOW THE PREVIOUS USER IN THE USER LIST SPREADSHEET!");
						System.out.println("What row should this new user occupy? (Must be 1 or greater)");
						inpt = sc.nextLine();
						if (!inpt.contains("cancel") && Integer.parseInt(inpt) > 0) {
							int newcell = Integer.parseInt(inpt);
							CreateUser(newname, newuuid, newcell);
						}
					}
				}
			}
			if (inpt.contains("stop")) {
				stop = true;
			}
		}
	}

	public static void CreateUser(String cName, String cUUID, int cCell) {
		Workbook wb = null;
		try {
			FileInputStream fis = new FileInputStream(wbLoc);
			wb = new XSSFWorkbook(fis);
		} catch (IOException e1) {
			e1.printStackTrace();
		}
		if (wb != null) {
			Sheet sheet = wb.getSheetAt(0);
			Row row = sheet.createRow(cCell - 1);
			Cell cell = row.createCell(0);
			cell.setCellValue(cUUID);
			cell = row.createCell(1);
			cell.setCellValue(cName);
			try {
				FileOutputStream out = new FileOutputStream(wbLoc);
				wb.write(out);
				out.close();
			} catch (Exception e) {
				e.printStackTrace();
			}
			System.out.println("The user was added successfully!");
		} else {
			System.out.println("The user list spreadsheet file cannot be null.");
		}
	}
}