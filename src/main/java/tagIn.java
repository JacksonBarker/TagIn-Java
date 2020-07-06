import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.FileNotFoundException;
import java.io.InputStream;
import java.io.InputStreamReader;
import java.text.DateFormat;
import java.text.SimpleDateFormat;
import java.util.*;
import java.security.GeneralSecurityException;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import com.google.api.client.auth.oauth2.Credential;
import com.google.api.client.extensions.java6.auth.oauth2.AuthorizationCodeInstalledApp;
import com.google.api.client.extensions.jetty.auth.oauth2.LocalServerReceiver;
import com.google.api.client.googleapis.auth.oauth2.GoogleAuthorizationCodeFlow;
import com.google.api.client.googleapis.auth.oauth2.GoogleClientSecrets;
import com.google.api.client.googleapis.javanet.GoogleNetHttpTransport;
import com.google.api.client.http.javanet.NetHttpTransport;
import com.google.api.client.json.JsonFactory;
import com.google.api.client.json.jackson2.JacksonFactory;
import com.google.api.client.util.store.FileDataStoreFactory;
import com.google.api.services.sheets.v4.Sheets;
import com.google.api.services.sheets.v4.SheetsScopes;
import com.google.api.services.sheets.v4.model.ValueRange;

public class tagIn {

	public static String inpt;
	public static String wbLoc = "C:/TagIn-Java/src/main/resources/userlist.xlsx";
	public static boolean stop = false;

	private static final String APPLICATION_NAME = "Google Sheets API Java Quickstart";
	private static final JsonFactory JSON_FACTORY = JacksonFactory.getDefaultInstance();
	private static final String TOKENS_DIRECTORY_PATH = "tokens";

	private static final List<String> SCOPES = Collections.singletonList(SheetsScopes.SPREADSHEETS_READONLY);
	private static final String CREDENTIALS_FILE_PATH = "/credentials.json";

	private static Credential getCredentials(final NetHttpTransport HTTP_TRANSPORT) throws IOException {
		InputStream in = tagIn.class.getResourceAsStream(CREDENTIALS_FILE_PATH);
		if (in == null) {
			throw new FileNotFoundException("Resource not found: " + CREDENTIALS_FILE_PATH);
		}
		GoogleClientSecrets clientSecrets = GoogleClientSecrets.load(JSON_FACTORY, new InputStreamReader(in));

		GoogleAuthorizationCodeFlow flow = new GoogleAuthorizationCodeFlow.Builder(
				HTTP_TRANSPORT, JSON_FACTORY, clientSecrets, SCOPES)
				.setDataStoreFactory(new FileDataStoreFactory(new java.io.File(TOKENS_DIRECTORY_PATH)))
				.setAccessType("offline")
				.build();
		LocalServerReceiver receiver = new LocalServerReceiver.Builder().setPort(8888).build();
		return new AuthorizationCodeInstalledApp(flow, receiver).authorize("user");
	}

	public static void main(String[] args) throws IOException, GeneralSecurityException {
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
			} else if (inpt.contains("newday")) {
				System.out.println("Are you sure you would like to start a new day?");
				System.out.println("Type 'confirm' to start new day.");
				inpt = sc.nextLine();
				if (inpt.contains("confirm")) {
					NewDay();
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
		boolean Found = false;
		boolean EndOfList = false;
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
				try {
					Cell cell = row.getCell(method);
					if (info.equals(cell.getStringCellValue())) {
						cell = row.getCell(0);
						System.out.println("The user's UUID is: " + cell.getStringCellValue());
						cell = row.getCell(1);
						System.out.println("The user's name is: " + cell.getStringCellValue());
						System.out.println("The user's row is: " + (i + 1));
						cell = row.getCell(2);
						System.out.println("Is user signed in?: " + cell.getBooleanCellValue());

						Found = true;
					}
				} catch (NullPointerException e2) {
					Found = true;
					EndOfList = true;
				}
			}
		} else {
			System.out.println("The user list spreadsheet file cannot be null.");
		}
			if (EndOfList) {
				System.out.println("A user matching the information provided was not found");
			}

		}

		public static void NewDay() throws IOException, GeneralSecurityException {
			final NetHttpTransport HTTP_TRANSPORT = GoogleNetHttpTransport.newTrustedTransport();
			final String spreadsheetId = "1WxijUbdukzbkBVXbWcBUzH5aOo-Cy9mX6pixO710EMU";
			final String range = "A4:A";

			DateFormat dateFormat = new SimpleDateFormat("MM/dd");
			Date date = new Date();
			String dateString = dateFormat.format(date);

			Sheets service = new Sheets.Builder(HTTP_TRANSPORT, JSON_FACTORY, getCredentials(HTTP_TRANSPORT))
					.setApplicationName(APPLICATION_NAME)
					.build();
			ValueRange response = service.spreadsheets().values()
					.get(spreadsheetId, range)
					.execute();
			List<List<Object>> values = response.getValues();
			if (values == null || values.isEmpty()) {
				System.out.println("No data found.");
			} else {
				for (List row : values) {
					System.out.printf("%s\n", row.get(0));
				}
			}
		}
	}