package tagIn;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.text.DateFormat;
import java.text.SimpleDateFormat;
import java.util.Date;
import java.util.Scanner;
import java.util.UUID;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class tagIn {

	public static String inpt;
	public static String wbLoc = "C:/eclipse/TagIn-Java/demosheet.xlsx";
	public static Boolean stop = false;
	public static DateFormat dateFormat = new SimpleDateFormat("hh:mm aa");

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
					System.out.println("WARNING PLEASE CHECK THAT THE ROW DOESN'T CONTAIN ANOTHER USER AND IS DIRECTLY BELOW THE PREVIOUS USER IN THE USER LIST SPREADSHEET!");
					System.out.println("What row should this new user occupy? (Must be 1 or greater)");
					inpt = sc.nextLine();
					if (!inpt.contains("cancel") && Integer.parseInt(inpt) > 0) {
						int newcell = Integer.parseInt(inpt);
						CreateUser(newname, newcell);
					}
				}
			} else if (inpt.contains("find")) {
				System.out.println("What are you using to find the user's information? (Name or UUID)");
				System.out.println("Send 'cancel' at any time to cancel user search.");
				inpt = sc.nextLine();
				if (!inpt.contains("cancel")) {
					if (inpt.toLowerCase().contains("uuid")) {
						System.out.println("What is the user's UUID?");
						inpt = sc.nextLine();
						if (!inpt.contains("cancel")) {
							FindUser(0, inpt);
						}
					} else if (inpt.toLowerCase().contains("name")) {
						System.out.println("What is the user's name?");
						inpt = sc.nextLine();
						if (!inpt.contains("cancel"))
							FindUser(1, inpt);
					}
				}
			} else if (inpt.contains("stop")) {
				stop = true;
			}

		}
	}

	public static void CreateUser(String cName, int cCell) {
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
			UUID uuid = UUID.randomUUID();
			cell.setCellValue(uuid.toString());
			cell = row.createCell(1);
			cell.setCellValue(cName);
			cell = row.createCell(2);
			cell.setCellType(CellType.BOOLEAN);
			cell.setCellValue(false);
			cell = row.createCell(3);
			dateFormat = new SimpleDateFormat("hh:mm aa");
			cell.setCellValue(dateFormat.format(new Date()).toString());

			try {
				FileOutputStream out = new FileOutputStream(wbLoc);
				wb.write(out);
				out.close();
			} catch (Exception e) {
				e.printStackTrace();
			}
			System.out.println("The user was generated with the UUID: " + uuid);

		} else {
			System.out.println("The user list spreadsheet file cannot be null.");
		}
	}

	public static void FindUser(int method, String info) {
		Boolean Found = false;
		Boolean EndOfList = false;
		Workbook wb = null;
		try {
			FileInputStream fis = new FileInputStream(wbLoc);
			wb = new XSSFWorkbook(fis);
		} catch (IOException e1) {
			e1.printStackTrace();
		}
		if (wb != null) {
			Sheet sheet = wb.getSheetAt(0);
			for (int i = 0; !Found; i++) {
				Row row = sheet.getRow(i);
				Cell cell = row.getCell(method);
				if (info.equals(cell.getStringCellValue())) {
					cell = row.getCell(0);
					System.out.println("The user's UUID is: " + cell.getStringCellValue());
					cell = row.getCell(1);
					System.out.println("The user's name is: " + cell.getStringCellValue());
					System.out.println("The user's row is: " + (i+1));
					cell = row.getCell(2);
					System.out.println("Is user signed in?: " + cell.getBooleanCellValue());
					cell = row.getCell(3);
					System.out.println("Last user sign-in/out: " + cell.getStringCellValue());

					Found = true;
				} else if (cell.getStringCellValue() == "") {
					EndOfList = true;
					Found = true;
				}
			}
		} else {
			System.out.println("The user list spreadsheet file cannot be null.");
		}
		if (EndOfList) {
			System.out.println("A user matching the information provided was not found");
		}

	}
}