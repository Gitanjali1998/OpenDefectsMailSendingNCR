package com.ncr.opendefects.mail;

import java.io.BufferedWriter;

import java.io.File;

import java.io.FileWriter;

import java.io.IOException;

import java.sql.Connection;

import java.sql.DriverManager;

import java.sql.ResultSet;

import java.sql.SQLException;

import java.sql.Statement;

import java.sql.Timestamp;

import java.text.SimpleDateFormat;

import java.util.ArrayList;

import java.util.Arrays;

import java.util.Calendar;

import java.util.Collections;

import java.util.Date;

import java.util.HashMap;

import java.util.HashSet;

import java.util.LinkedHashMap;

import java.util.List;

import java.util.Map;

import java.util.Properties;

import java.util.Set;

import javax.mail.Message;

import javax.mail.MessagingException;

import javax.mail.PasswordAuthentication;

import javax.mail.Session;

import javax.mail.Transport;

import javax.mail.internet.InternetAddress;

import javax.mail.internet.MimeMessage;

import org.apache.poi.hssf.util.HSSFColor;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;

import org.apache.poi.ss.usermodel.Cell;

import org.apache.poi.ss.usermodel.CellStyle;

import org.apache.poi.ss.usermodel.Color;

import org.apache.poi.ss.usermodel.CreationHelper;

import org.apache.poi.ss.usermodel.DataFormatter;

import org.apache.poi.ss.usermodel.DateUtil;

import org.apache.poi.ss.usermodel.FormulaEvaluator;

import org.apache.poi.ss.usermodel.Row;

import org.apache.poi.ss.usermodel.Sheet;

import org.apache.poi.ss.usermodel.Workbook;

import org.apache.poi.ss.usermodel.WorkbookFactory;

import org.apache.poi.ss.util.CellRangeAddress;

import org.apache.poi.xssf.usermodel.XSSFColor;
import java.sql.Timestamp;
import java.text.ParseException;
import java.text.SimpleDateFormat;
import java.util.*;

public class OpenDefectsMail {

	static Set<String> jiraOwners = new HashSet<String>();

	static List<Integer> colorIndexes = new ArrayList<Integer>();

	private static DataFormatter dataFormatter = new DataFormatter();

	static Map<String, List<Map<String, String>>> allOwnersMap = new LinkedHashMap<String, List<Map<String, String>>>();

	static Map<String, List<Map<String, String>>> assigneeNamesMap = new LinkedHashMap<String, List<Map<String, String>>>();

	static Map<String, List<Map<String, String>>> ownerQLIDMap = new LinkedHashMap<String, List<Map<String, String>>>();

	static Map<String, List<Map<String, String>>> assigneeQLIDMap = new LinkedHashMap<String, List<Map<String, String>>>();
	static Map<String,String>NotFoundNamesFromTeradataMap  = new LinkedHashMap<>();
	static List<String>NotFoundNamesFromTeradata=new ArrayList<>();

	public static void main(String[] args) throws SQLException, IOException {

		 String filename = "\\\\susday5469\\SW Quality Reports\\External Open Defects by Status.xlsx";

		 //String filename = "E:/External Open Defects by Status.xlsx";

		//String filename = "G:/openDefectMail/External Open Defects by Status1.xlsx";
		//String NotFoundNames="G:/NotFoundNames.xlsx";
		String AdaptiveFile = "\\susday5469\\Reports\\Adaptive Insights Pull";
		//String AdaptiveFile="G:/AdaptivePullISights";

		// E:\SW Quality Reports

		// String filename = "E:\\SW Quality Reports\\Book1.xlsx";

		String sheetname = "SLA Days";
		String NotFoundNamesSheet="Sheet1";

		OpenDefectsMail sendEmailObject = new OpenDefectsMail();
		if (sendEmailObject.checkFileModification(AdaptiveFile)) {
			String sub = "New Adaptive Insights Pull file is Updated";
			String msg = "<p>Hi Geethanjali and Sathish ,<br><br>New Adaptive Insights Pull file is Updated. So Please Update the new budget data into agilecraft.<br> File modification date is ";
			sendEmailObject.notifyUser(sub, msg);
		}
		// Connection conn=sendEmailObject.teraDataConnection();

		// System.out.println("QLID = "+sendEmailObject.getQLID(conn,"Sathish
		// Kodadhada"));

		// System.out.println("Name = "+getName(conn,"SK185620"));

		// Building the data

		List<String> columnNames = sendEmailObject.getColumns();
		//List<String> columnNamesForNotFoundNames = sendEmailObject.getColumnsForNotFoundNames();
		System.out.println("column names are " + columnNames);
		if (sendEmailObject.checkFileModification(filename)) {

			// sendEmailObject.buildOpenDefectData(columnNames);
			// System.out.println("adaptive file is modified");
			// exit(0);
			sendEmailObject.extractOpenDefectData(filename, sheetname, columnNames, "Assignee QLID","Jira Project Owner QLID");
			//sendEmailObject.extractOpenDefectDataForNotFoundNames(NotFoundNames, NotFoundNamesSheet);

			sendEmailObject.sendMail();
			System.out.println("not found names in teradata"+NotFoundNamesFromTeradataMap);
			System.out.println("all owners map is "+allOwnersMap);
			System.out.println("assignee names  map is "+assigneeNamesMap);
			System.out.println("assignee qlid  map is "+assigneeQLIDMap);
			System.out.println("owners  map is "+ownerQLIDMap);

			 System.out.println("ownerQLIDMap = "+ownerQLIDMap.keySet().size());

		     System.out.println("assigneeQLIDMap = "+assigneeQLIDMap.keySet().size());

			 System.out.println("mailToAllOwners = "+allOwnersMap.keySet().size());

			  System.out.println("assigneeNamesMap = "+assigneeNamesMap.keySet().size());

			// System.out.println("ownerQLIDMap = "+ownerQLIDMap.keySet());

			// System.out.println("assigneeQLIDMap = "+assigneeQLIDMap.keySet());

			// System.out.println("mailToAllOwners = "+allOwnersMap.keySet());

			//System.out.println("assigneeNamesMap = " + assigneeNamesMap.keySet());

			// for(String qlid:ownerQLIDMap.keySet())

			// System.out.println(qlid+" "+ownerQLIDMap.get(qlid).size());

			// sendEmailObject.extractOpenDefectData(filename, sheetname, columnNames,
			// "Assignee QLID", "Jira Project Owner QLID");

			// sendEmailObject.sendMail();

		}

		else {

			String sub = "External Open Defects File not Updated - ";
			String msg = "<p>Hi Geetha and Sathish ,<br><br> The External Open Defects Excel file in the SUSDAY5469 is not updated on ";
			sendEmailObject.notifyUser(sub, msg);

		}

	}

	private static void exit(int i) {
		// TODO Auto-generated method stub

	}

	private String getQLID(Connection conn, String name) throws SQLException {

		String qlid = null;

		// Connection conn=this.teraDataConnection();

		Statement stmt = conn.createStatement();

		String query = "\r\n" +

				"select distinct quick_look_id,common_name,display_name,update_date_time from vint.person_directory_vw where common_name = '"
				+ name + "' or display_name = '" + name + "' order by update_date_time desc;";

		ResultSet rs = stmt.executeQuery(query);

		if (rs.next()) {

			System.out.println("In if Returning Qlid " + rs.getString(2));

			qlid = rs.getString(1);

		}

		stmt.close();

		rs.close();

		// conn.close();

		return qlid;

	}

	private static String getName(Connection conn, String qlid) throws SQLException {

		String name = null;

		// System.out.println("In GetName -->");

		Statement st = conn.createStatement();

		// System.out.println("Statement created");

		ResultSet rs = st.executeQuery(
				"select quick_look_id,common_name from vint.person_directory_vw where quick_look_id='" + qlid + "'");

		// System.out.println("Result Set created");

		if (rs.next())

		{

			// System.out.println("In if Returning name "+rs.getString(2));

			name = rs.getString(2);

			// return name;

		}

		st.close();

		rs.close();

		// TODO Auto-generated method stub

		return name;

	}

	public Connection teraDataConnection() {

		String url = "jdbc:teradata://t61edw.daytonoh.ncr.com";

		String user = "t2pg185114";

		String pass = "sto751s9";

		try {

			// DriverManager.registerDriver(new oracle.jdbc.OracleDriver());

			try {

				Class.forName("com.teradata.jdbc.TeraDriver");

			} catch (ClassNotFoundException e) {

				// TODO Auto-generated catch block

				e.printStackTrace();

			}

			// Reference to connection interface

			Connection con = DriverManager.getConnection(url, user, pass);

			Statement st = con.createStatement();

			System.out.println("Conncetion Successfull");

			return con;

			/*
			 * ResultSet rs=st.
			 * executeQuery("select quick_look_id,common_name from vint.person_directory_vw where quick_look_id='SK185620'"
			 * );
			 * 
			 * if(rs.next())
			 * 
			 * System.out.println(rs.getString(1)+" "+rs.getString(2));
			 */

		}

		catch (Exception e) {

			e.printStackTrace();

		}

		return null;

	}

	private List<String> getColumns() {

		List<String> columnNames = new ArrayList<String>();

		columnNames.add("Product");

		columnNames.add("Issue Key");

		columnNames.add("Summary");

		columnNames.add("Status");

		columnNames.add("Customer Name");

		columnNames.add("Severity");

		columnNames.add("Priority");

		columnNames.add("Created");

		columnNames.add("Resolved Date");

		columnNames.add("Sla Measure");

		columnNames.add("SLA Days Open vs Target");

		columnNames.add("SLA Initial Response (Days in Triage) vs Target");

		columnNames.add("Update Freq Days Since Last Update");

		columnNames.add("Assignee");

		columnNames.add("Assignee QLID");

		columnNames.add("Jira Project Owner QLID");
		columnNames.add("CFNS Update Date");
		columnNames.add("CFNS Update");
		return columnNames;

	}
	private List<String> getColumnsForNotFoundNames(){
		List<String> columnNames1 = new ArrayList<String>();

		columnNames1.add("AssigneeName");

		columnNames1.add("Qlid");

		
		return columnNames1;
		
	}

	private boolean checkFileModification(String filename) {
		System.out.println("checking for modificaton ");
		File file = new File(filename);

		if (file.exists())

		{
			System.out.println("file is existed ");
			SimpleDateFormat sdf = new SimpleDateFormat("MM/dd/yyyy");

			String lastModifiedDate = sdf.format(file.lastModified());

			System.out.println("lastModifiedDate = " + sdf.format(file.lastModified()));

			Date today = new Date();

			// Date before=new Date(f.lastModified());

			System.out.println("today = " + sdf.format(today));
			System.out.println("last modified = " + lastModifiedDate);
			String present = sdf.format(today);

			if (lastModifiedDate.equals(present))

			{

				return true;

			}

			else {

				return false;

			}

		}

		return false;

	}

	/*
	 * private void buildOpenDefectData(List<String> columnNames) {
	 * 
	 * 
	 * 
	 * 
	 * 
	 * 
	 * 
	 * // return opendefectdata;
	 * 
	 * }
	 */

	// This function constructs data according to the excel sheet
	
	public void extractOpenDefectDataForNotFoundNames(String excelFilepath, String sheetName) {
		Sheet sheet = getSheet(excelFilepath, sheetName);
		if (sheet == null) {

			return;

		}

		String id = null;
		for (Row row : sheet) {
			if (row.getRowNum() == 0) {
				continue;
			}
			String name = getCellValue(row.getCell(0));
			String qlid = getCellValue(row.getCell(1));
			NotFoundNamesFromTeradataMap.put(name, qlid);
			
			
		}
		
	}
	public void extractOpenDefectData(String excelFilepath, String sheetName, List<String> columnNames,

			String assigneeId, String ownerID) {

		// Fetch the sheet at the specified sheet position.

		Sheet sheet = getSheet(excelFilepath, sheetName);

		// If the sheet is not present, there is nothing to do

		if (sheet == null) {

			return;

		}

		String id = null;

		Map<String, String> openDefectDetails = null;

		Map<String, Integer> columnNameWithIndex = null;

		Map<String, List<Map<String, String>>> openDefectInfo = new LinkedHashMap<String, List<Map<String, String>>>();

		int rowCount = 0;

		// Iterate over each row in the sheet

		for (Row row : sheet) {

			// Skip the first row (i.e. heading row) and populate the column name and column

			// number mapping if its not already

			// populated

			if (row.getRowNum() == 0) {

				columnNameWithIndex = getColumnNumbers(row, columnNames);

				// System.out.println("columns " + columnNameWithIndex);

				colorIndexes.add(columnNameWithIndex.get("SLA Initial Response (Days in Triage) vs Target"));

				colorIndexes.add(columnNameWithIndex.get("SLA Days Open vs Target"));

				colorIndexes.add(columnNameWithIndex.get("Update Freq Days Since Last Update"));

				Collections.sort(colorIndexes);

				// System.out.println("colors " + colorIndexes);

				// If column numbers cannot be determined then there is nothing to do

				if (columnNameWithIndex == null) {

					return;

				}

				continue;

			}

			// boolean isMobile = false;

			// System.out.println(rowCount++);

			id = getCellValue(row.getCell(columnNameWithIndex.get(assigneeId)));

			if (id.contentEquals("")) {

				// System.out.println("Inside AQLID null");

				id = getCellValue(row.getCell(columnNameWithIndex.get("Assignee")));

				if (id.contentEquals("")) {

					// System.out.println("Inside Name null");

					id = getCellValue(row.getCell(columnNameWithIndex.get(ownerID)));

					putData(ownerQLIDMap, id, row, columnNameWithIndex, columnNames);

				}

				else {

					// System.out.println("Inside Name");

					putData(assigneeNamesMap, id, row, columnNameWithIndex, columnNames);

				}

			}

			else {

				 System.out.println(id);

				putData(assigneeQLIDMap, id, row, columnNameWithIndex, columnNames);

			}

			// System.out.println("Owner");

			String ownerQLID = getCellValue(row.getCell(columnNameWithIndex.get(ownerID)));

			putData(allOwnersMap, ownerQLID, row, columnNameWithIndex, columnNames);

			/*
			 * boolean flag = false;
			 * 
			 * id = getCellValue(row.getCell(columnNameWithIndex.get(idColumnName)));
			 * 
			 * if (id.contentEquals("")) {
			 * 
			 * id = getCellValue(row.getCell(columnNameWithIndex.get(ownerID)));
			 * 
			 * flag = true;
			 * 
			 * }
			 * 
			 * if (flag)
			 * 
			 * putData(ownerQLID, id, row, columnNameWithIndex, columnNames);
			 * 
			 * else
			 * 
			 * putData(assigneeQLID, id, row, columnNameWithIndex, columnNames);
			 */

			// break;

		}

		sheet = null;

	}

	// Returns sheet in the excel file

	private static Sheet getSheet(String excelFilepath, String sheetName) {

		// If the workbook is not specified there is nothing to do

		if (excelFilepath == null || excelFilepath.length() == 0) {

			return null;

		}

		// Access the workbook. The workbook must exist and be readable.

		Workbook projectDataWorkbook = null;

		try {

			projectDataWorkbook = WorkbookFactory.create(new File(excelFilepath));

		} catch (InvalidFormatException e) {

			e.getMessage();

		} catch (IOException e) {

			e.getMessage();

		}

		// If the workbook cannot be fetched there is nothing to do

		if (projectDataWorkbook == null) {

			return null;

		}

		// Fetch the sheet at the specified sheet position.

		Sheet sheet = null;

		if (sheetName.equals("")) {

			sheet = projectDataWorkbook.getSheetAt(0);

		} else {

			sheet = projectDataWorkbook.getSheet(sheetName);

		}

		return sheet;

	}

	// This returns cloumn names mapped by their position(i.e column number) in

	// excel file

	private static Map<String, Integer> getColumnNumbers(Row row, List<String> columnNames) {

		// A Map of column names and numbers. Column Name forms the key whereas the

		// column number forms the value.

		Map<String, Integer> columnNumbers = new HashMap<String, Integer>();

		// Get the minimum and maximum column indexes of the row

		short minColIndex = row.getFirstCellNum();

		short maxColIndex = row.getLastCellNum();

		// Iterate over the column index range

		for (int colIndex = minColIndex; colIndex < maxColIndex; colIndex++) {

			// Fetch the cell for the specified column index

			Cell cell = row.getCell(colIndex);

			// If the cell is null, we ignore it.

			if (cell == null) {

				continue;

			}

			// Get the cell value which is actually the column name

			String columnName = cell.getStringCellValue();

			// If the column name is in the list of column names then we record the column

			// index for that column name

			if (columnNames.contains(columnName)) {

				columnNumbers.put(columnName, Integer.valueOf(colIndex));

			}

		}

		return columnNumbers;

	}

	public static void putData(Map<String, List<Map<String, String>>> openDefectInfo, String id, Row row,

			Map<String, Integer> columnNameWithIndex, List<String> columnNames) {

		if (!openDefectInfo.containsKey(id)) {

			List<Map<String, String>> list = new ArrayList<Map<String, String>>();

			HashMap<String, String> openDefectDetails = new HashMap<String, String>();

			for (String columnName : columnNames) {

				Integer cellnum = columnNameWithIndex.get(columnName);
				System.out.println("column  name "+columnName);
				System.out.println("cell num"+cellnum);
				Cell cell = row.getCell(cellnum);

				openDefectDetails.put(columnName, getCellValue(cell));

				if (colorIndexes.contains(cellnum)) {

					try {

						CellStyle cellStyle = cell.getCellStyle();

						Color color = cellStyle.getFillForegroundColorColor();

						if (color != null) {

							if (color instanceof XSSFColor) {

								// System.out.println("in if");

								openDefectDetails.put(String.valueOf(cellnum), ((XSSFColor) color).getARGBHex());

								// System.out.println(((XSSFColor) color).getARGBHex());

							} else if (color instanceof HSSFColor) {

								if (!(color instanceof HSSFColor.AUTOMATIC))

									System.out.println(((HSSFColor) color).getHexString());

							}

						}

					} catch (Exception e) {

					}

				}

			}

			list.add(openDefectDetails);

			openDefectInfo.put(id, list);

			// break;

		} else {

			List<Map<String, String>> list = openDefectInfo.get(id);

			HashMap<String, String> openDefectDetails = new HashMap<String, String>();

			// openDefectDetails.put("isMobile", Boolean.toString(isMobile));

			for (String columnName : columnNames) {

				Integer cellnum = columnNameWithIndex.get(columnName);

				Cell cell = row.getCell(cellnum);

				openDefectDetails.put(columnName, getCellValue(cell));

				// This fetches colors of cell in excel for columns M(12),N(13) and O(14)

				if (colorIndexes.contains(cellnum)) {

					// System.out.println(id);

					try {

						CellStyle cellStyle = cell.getCellStyle();

						Color color = cellStyle.getFillForegroundColorColor();

						if (color != null) {

							if (color instanceof XSSFColor) {

								// System.out.println("in if");

								openDefectDetails.put(String.valueOf(cellnum), ((XSSFColor) color).getARGBHex());

								// System.out.println(((XSSFColor) color).getARGBHex());

							} else if (color instanceof HSSFColor) {

								if (!(color instanceof HSSFColor.AUTOMATIC))

									System.out.println(((HSSFColor) color).getHexString());

							}

						}

					} catch (Exception e) {

					}

				}

			}

			list.add(openDefectDetails);

			openDefectInfo.put(id, list);

		}

	}

	// This returns data in the cell

	private static String getCellValue(Cell cell) {

		String data = "";

		if (cell == null) {

			return data;

		}

		switch (cell.getCellType()) {

		case Cell.CELL_TYPE_STRING:

			data = cell.getRichStringCellValue().getString();

			break;

		case Cell.CELL_TYPE_NUMERIC:

			// if (DateUtil.isCellDateFormatted(cell)) {

			// SimpleDateFormat dateFormat = new SimpleDateFormat("yyyy-MM-dd");

			// data = dateFormat.format(cell.getDateCellValue());

			// } else {

			// data = String.valueOf(cell.getNumericCellValue());

			data = dataFormatter.formatCellValue(cell);

			// }

			break;

		case Cell.CELL_TYPE_BOOLEAN:

			data = String.valueOf(cell.getBooleanCellValue());

			break;

		case Cell.CELL_TYPE_FORMULA:

			Workbook wb = cell.getSheet().getWorkbook();

			CreationHelper createHelper = wb.getCreationHelper();

			FormulaEvaluator evaluator = createHelper.createFormulaEvaluator();

			switch (evaluator.evaluateFormulaCell(cell)) {

			case Cell.CELL_TYPE_BOOLEAN:

				data = String.valueOf(cell.getBooleanCellValue());

				break;

			case Cell.CELL_TYPE_NUMERIC:

				if (DateUtil.isCellDateFormatted(cell)) {

					// if (cell.getDateCellValue() == null) {

					// data = null;

					// } else {

					// if (LOG.isInfoEnabled()) {

					// LOG.info("Found Date Cell " + cell.getRichStringCellValue().getString());

					// }

					SimpleDateFormat dateFormat = new SimpleDateFormat("yyyy-MM-dd");

					data = dateFormat.format(cell.getDateCellValue());

				} else {

					// data = String.valueOf(cell.getNumericCellValue());

					data = dataFormatter.formatCellValue(cell);

				}

				break;

			case Cell.CELL_TYPE_STRING:

				data = cell.getRichStringCellValue().getString();

				break;

			case Cell.CELL_TYPE_BLANK:

				data = null;

				break;

			case Cell.CELL_TYPE_ERROR:

				System.out.println(cell.getErrorCellValue());

				break;

			// data = cell.getCellFormula();

			// break;

			default:

				data = cell.getStringCellValue();

			}

		}

		return data.trim();

	}

	private void notifyUser(String sub, String msg) {

		final String username = "da230140@ncr.com";

		final String password = "mhFQpnqrTNvH36OM42#l?GOF7PsUJO";

		String cc = "sk185620@ncr.com";

		Properties props = new Properties();

		props.put("mail.smtp.auth", "true");

		props.put("mail.smtp.host", "ncrusout1.ncr.com");

		props.put("mail.smtp.port", "25");

		Session session = Session.getInstance(props, new javax.mail.Authenticator() {

			protected PasswordAuthentication getPasswordAuthentication() {

				return new PasswordAuthentication(username, password);

			}

		});

		try {

			SimpleDateFormat sdf = new SimpleDateFormat("MM/dd/yyyy");

			Date today = new Date();

			MimeMessage message = new MimeMessage(session);

			message.setFrom(new InternetAddress("dg230087@ncr.com"));

			// message.setRecipients(Message.RecipientType.TO,
			// InternetAddress.parse("AD230173@ncr.com"));
			message.setRecipients(Message.RecipientType.TO, InternetAddress.parse("gj185048@ncr.com,sk185620@ncr.com"));

			message.addRecipients(Message.RecipientType.CC, cc);

			// message.setSubject("External Open Defects File not Updated - ");
			message.setSubject(sub);
			message.setContent(msg + sdf.format(today) + ".</p>", "text/html");
			/*
			 * message.setContent(
			 * "<p>Hi Prasanth,<br><br> The External Open Defects Excel file in the SUSDAY5469 is not updated on "
			 * + sdf.format(today) + ".</p>",
			 * 
			 * "text/html");
			 */

			Transport.send(message);

			System.out.println("Done");

			// break;

		} catch (Exception e) {

			// e.printStackTrace();

		}

	}

	// Mail sender code
	
	

	public void sendMail() throws IOException {

		final String username = "da230140@ncr.com";

		final String password = "mhFQpnqrTNvH36OM42#l?GOF7PsUJO";

		Connection conn = this.teraDataConnection();

		String cc = "gj185048@ncr.com,sk185620@ncr.com";

		Properties props = new Properties();

		props.put("mail.smtp.auth", "true");

		props.put("mail.smtp.host", "ncrusout1.ncr.com");

		props.put("mail.smtp.port", "25");

		Session session = Session.getInstance(props, new javax.mail.Authenticator() {

			protected PasswordAuthentication getPasswordAuthentication() {

				return new PasswordAuthentication(username, password);

			}

		});

		for (String qlid : ownerQLIDMap.keySet()) {

			try {

				MimeMessage message = new MimeMessage(session);

				message.setFrom(new InternetAddress("dg230087@ncr.com"));

				// Map<String, String> temporary = projectOwners.get(qlid).get(0);

				message.setRecipients(Message.RecipientType.TO, InternetAddress.parse(qlid + "@ncr.com"));

				message.addRecipients(Message.RecipientType.CC, cc);

				System.out.println("Sending to " + qlid);

				//String name = this.getName(conn, qlid);

				//System.out.println("name = " + name);

				// System.out.println("name = "+name);

				String ownerName = qlid;

				//if (name != null)

					//ownerName = name;

				message.setSubject("Jira External Open Defects Notification - " + ownerName);

				String summary = "";

				summary = "<p>Dear " + ownerName

						+ ",<br><br>The following Externally found Defects are open and unassigned on a JIRA project you own.   The SLA information is provided here to help you manage your open items against the SLAs. Please see legend below for more information on actions to be taken.</p>";

				String bottomLine = "<br><p style=\"color:green\">For more information on the SLA KPIs please see <a href="

						+ "https://confluence.ncr.com/display/SWD/SW+Quality+Metrics+Definitions+and+Targets+by+Team"

						+ ">" + "Metrics definition Confluence Page" + "</a></p>";

				List<Map<String, String>> qlidInfo = ownerQLIDMap.get(qlid);

				message.setContent(summary + getBody(qlid, qlidInfo,0) + "<br><br>" + getLegend() + bottomLine,

						"text/html");

				Transport.send(message);

				System.out.println("Done");

				// break;

			} catch (Exception e) {

				// e.printStackTrace();

			}

		}

		for (String qlid : assigneeQLIDMap.keySet()) {

			try {

				MimeMessage message = new MimeMessage(session);

				message.setFrom(new InternetAddress("dg230087@ncr.com"));

				// Map<String, String> temporary = assignees.get(qlid).get(0);

				message.setRecipients(Message.RecipientType.TO, InternetAddress.parse(qlid + "@ncr.com"));

				message.addRecipients(Message.RecipientType.CC, cc);

				System.out.println("Sending to " + qlid);

				String name = this.getName(conn, qlid);

				System.out.println("name = " + name);

				// System.out.println("name = "+name);

				String assigneeName = qlid;

				if (name != null)

					assigneeName = name;

				message.setSubject("Jira External Open Defects Notification - " + assigneeName);

				// System.out.println("Sending to " + qlid);

				String summary = "";

				summary = "<p>Dear " + assigneeName + ","

						+ "<br><br>The following Externally found Defects are open with you as the assignee.  "

						+ "The SLA information is provided here to help you manage your open items against the SLAs. "

						+ "Please see legend below for more information on actions to be taken.</p><br>";

				String bottomLine = "<br><p style=\"color:green\">For more information on the SLA KPIs please see <a href="

						+ "https://confluence.ncr.com/display/SWD/SW+Quality+Metrics+Definitions+and+Targets+by+Team"

						+ ">" + "Metrics definition Confluence Page" + "</a></p>";

				List<Map<String, String>> qlidInfo = assigneeQLIDMap.get(qlid);

				message.setContent(summary + getBody(qlid, qlidInfo,0) + "<br><br>" + getLegend() + bottomLine,

						"text/html");

				Transport.send(message);

				System.out.println("Done");

				// break;

			} catch (Exception e) {

				// e.printStackTrace();

			}

		}

		//Date date = new Date();

		//SimpleDateFormat dateFormat = new SimpleDateFormat("yyyy-MM-dd HH-mm-ss");

		//File f = new File(dateFormat.format(date) + " - Exception.txt");

		//BufferedWriter bw = new BufferedWriter(new FileWriter(f));

		System.out.println("===================================================================== assignee names");
		//bw.write("mails sending based n assignee names "+ "\n");

		/*for (String assigneeName : assigneeNamesMap.keySet()) {

			try {

				System.out.println(assigneeName);

				MimeMessage message = new MimeMessage(session);

				message.setFrom(new InternetAddress("dg230087@ncr.com"));

				// Map<String, String> temporary = projectOwners.get(qlid).get(0);

				String qlid = getQLID(conn, assigneeName);

				System.out.println("Sent to getName");

				if (qlid == null)

				{

					
                    
					//bw.write(assigneeName + " is not Found\n");
					qlid = NotFoundNamesFromTeradataMap.get(assigneeName);
					if (qlid == null) {
					NotFoundNamesFromTeradata.add(assigneeName);
					System.out.println("QLID is Not found for " + assigneeName);
					continue;
					}

				}

				//bw.write(assigneeName + " ==== " + qlid + "\n");

				message.setRecipients(Message.RecipientType.TO, InternetAddress.parse(qlid + "@ncr.com"));

				message.addRecipients(Message.RecipientType.CC, cc);

				System.out.println("Sending to " + qlid);

				// String name=this.getName(conn, qlid);

				System.out.println("name = " + assigneeName);

				// System.out.println("name = "+name);

				message.setSubject("Jira External Open Defects Notification - " + assigneeName);

				String summary = "";

				summary = "<p>Dear " + assigneeName + ","

						+ "<br><br>The following Externally found Defects are open with you as the assignee.  "

						+ "The SLA information is provided here to help you manage your open items against the SLAs. "

						+ "Please see legend below for more information on actions to be taken.</p><br>";

				String bottomLine = "<br><p style=\"color:green\">For more information on the SLA KPIs please see <a href="

						+ "https://confluence.ncr.com/display/SWD/SW+Quality+Metrics+Definitions+and+Targets+by+Team"

						+ ">" + "Metrics definition Confluence Page" + "</a></p>";

				List<Map<String, String>> qlidInfo = assigneeNamesMap.get(assigneeName);

				message.setContent(summary + getBody(qlid, qlidInfo,0) + "<br><br>" + getLegend() + bottomLine,

						"text/html");

				//Transport.send(message);

				System.out.println("Done");

				// break;

			} catch (Exception e) {

				// e.printStackTrace();

			}

		}*/
		System.out.println("===================================================================== all owners");
		Calendar cal = Calendar.getInstance();
		int day = cal.get(Calendar.DAY_OF_WEEK);
		System.out.print("Today is " + day);
		// System.out.println("hiiiii ");
		for (String qlid : allOwnersMap.keySet()) {

			try {

				MimeMessage message = new MimeMessage(session);

				message.setFrom(new InternetAddress("dg230087@ncr.com"));

				// Map<String, String> temporary = projectOwners.get(qlid).get(0);

				message.setRecipients(Message.RecipientType.TO, InternetAddress.parse(qlid + "@ncr.com"));

				message.addRecipients(Message.RecipientType.CC, cc);

				System.out.println("Sending to " + qlid);

				String name = this.getName(conn, qlid);

				System.out.println("name = " + name);

				// System.out.println("name = "+name);

				String ownerName = qlid;

				if (name != null)

					ownerName = name;
				if (day == 2) {
				message.setSubject("Jira External Open Defects Notification - " + ownerName);
				}
				else {
					message.setSubject("Jira External Open S1/P1 Defects Notification -  " + ownerName);
				}

				String summary = "";

				summary = "<p>Dear " + ownerName

						+ ",<br><br>The following Externally found Defects are open on a Jira project on which you are the Project Lead. The SLA information is provided here to help you manage your teams open items against the SLAs. Please see legend below for more information on actions to be taken.</p>";

				String bottomLine = "<br><p style=\"color:green\">For more information on the SLA KPIs please see <a href="

						+ "https://confluence.ncr.com/display/SWD/SW+Quality+Metrics+Definitions+and+Targets+by+Team"

						+ ">" + "Metrics definition Confluence Page" + "</a></p>";

				List<Map<String, String>> qlidInfo = new ArrayList<>();
				List<Map<String, String>> highPriority = allOwnersMap.get(qlid);
				// allOwnersMap.get(qlid);
				if (day != 2) {
					// List<Map<String, String>> highPriority = new ArrayList<>();
					for (int i = 0; i < highPriority.size(); i++) {

						Map<String, String> mapOfqlidInfo = highPriority.get(i);
						if (mapOfqlidInfo.get("Severity").contentEquals("S1")
								|| mapOfqlidInfo.get("Priority").contentEquals("P1")) {
							qlidInfo.add(mapOfqlidInfo);
						}
					}
					

				} else {
					qlidInfo = allOwnersMap.get(qlid);
				}
               
				message.setContent(summary + getBody(qlid, qlidInfo,1) + "<br><br>" + getLegend() + bottomLine,

						"text/html");
                
			   if(qlidInfo.size()>0) {
				Transport.send(message);
			   }

				System.out.println("Done");

				// break;

			} catch (Exception e) {

				e.printStackTrace();

			}

		}
		for(int notFound = 0; notFound < NotFoundNamesFromTeradata.size() ; notFound++) {
			//bw.write(NotFoundNamesFromTeradata.get(notFound) + "  is not Found\n");     
		}
	
		
		
		//bw.close();
	}
	// In this function we are constructing body for each assignes's defects

	private String getBody(String quickLookID, List<Map<String, String>> qlidInfo,int extraColumn) {
		String mainBody = "<html xmlns:o=\"urn:schemas-microsoft-com:office:office\"\r\n"
				+ "xmlns:x=\"urn:schemas-microsoft-com:office:excel\"\r\n"
				+ "xmlns=\"http://www.w3.org/TR/REC-html40\">\r\n" + "\r\n" + "<head>\r\n"
				+ "<meta http-equiv=Content-Type content=\"text/html; charset=windows-1252\">\r\n"
				+ "<meta name=ProgId content=Excel.Sheet>\r\n"
				+ "<meta name=Generator content=\"Microsoft Excel 15\">\r\n"
				+ "<link rel=File-List href=\"OpenSummary_files/filelist.xml\">\r\n"
				+ "<style id=\"Temp_Report2_Styles\">\r\n" + "<!--table\r\n"
				+ " {mso-displayed-decimal-separator:\"\\.\";\r\n" + " mso-displayed-thousand-separator:\"\\,\";}\r\n"
				+ ".xl151886\r\n" + " {padding-top:1px;\r\n" + " padding-right:1px;\r\n" + " padding-left:1px;\r\n"
				+ " mso-ignore:padding;\r\n" + " color:black;\r\n" + " font-size:10.0pt;\r\n" + " font-weight:400;\r\n"
				+ " font-style:normal;\r\n" + " text-decoration:none;\r\n" + " font-family:Arial;\r\n"
				+ " mso-generic-font-family:auto;\r\n" + " mso-font-charset:0;\r\n" + " mso-number-format:General;\r\n"
				+ " text-align:general;\r\n" + " vertical-align:bottom;\r\n" + " mso-background-source:auto;\r\n"
				+ " mso-pattern:auto;\r\n" + " white-space:nowrap;}\r\n" + ".xl631886\r\n" + " {padding-top:1px;\r\n"
				+ " padding-right:1px;\r\n" + " padding-left:1px;\r\n" + " mso-ignore:padding;\r\n"
				+ " color:#333333;\r\n" + " font-size:9.0pt;\r\n" + " font-weight:400;\r\n" + " font-style:normal;\r\n"
				+ " text-decoration:none;\r\n" + " font-family:Arial;\r\n" + " mso-generic-font-family:auto;\r\n"
				+ " mso-font-charset:0;\r\n" + " mso-number-format:\"\\@\";\r\n" + " text-align:center;\r\n"
				+ " vertical-align:bottom;\r\n" + " border:.5pt solid #EBEBEB;\r\n" + " background:#F8FBFC;\r\n"
				+ " mso-pattern:white none;\r\n" + " white-space:nowrap;}\r\n" + ".xl641886\r\n"
				+ " {padding-top:1px;\r\n" + " padding-right:1px;\r\n" + " padding-left:1px;\r\n"
				+ " mso-ignore:padding;\r\n" + " color:#333333;\r\n" + " font-size:9.0pt;\r\n" + " font-weight:400;\r\n"
				+ " font-style:normal;\r\n" + " text-decoration:none;\r\n" + " font-family:Arial;\r\n"
				+ " mso-generic-font-family:auto;\r\n" + " mso-font-charset:0;\r\n" + " mso-number-format:General;\r\n"
				+ " text-align:right;\r\n" + " vertical-align:bottom;\r\n" + " border:.5pt solid #EBEBEB;\r\n"
				+ " background:#F8FBFC;\r\n" + " mso-pattern:white none;\r\n" + " white-space:nowrap;}\r\n"
				+ ".xl651886\r\n" + " {padding-top:1px;\r\n" + " padding-right:1px;\r\n" + " padding-left:1px;\r\n"
				+ " mso-ignore:padding;\r\n" + " color:#333333;\r\n" + " font-size:9.0pt;\r\n" + " font-weight:700;\r\n"
				+ " font-style:normal;\r\n" + " text-decoration:none;\r\n" + " font-family:Arial;\r\n"
				+ " mso-generic-font-family:auto;\r\n" + " mso-font-charset:0;\r\n" + " mso-number-format:General;\r\n"
				+ " text-align:right;\r\n" + " vertical-align:bottom;\r\n" + " border-top:.5pt solid #CAC9D9;\r\n"
				+ " border-right:.5pt solid #EBEBEB;\r\n" + " border-bottom:.5pt solid #EBEBEB;\r\n"
				+ " border-left:.5pt solid #EBEBEB;\r\n" + " background:white;\r\n" + " mso-pattern:white none;\r\n"
				+ " white-space:nowrap;}\r\n" + ".xl661886\r\n" + " {padding-top:1px;\r\n" + " padding-right:1px;\r\n"
				+ " padding-left:1px;\r\n" + " mso-ignore:padding;\r\n" + " color:#333333;\r\n"
				+ " font-size:9.0pt;\r\n" + " font-weight:700;\r\n" + " font-style:normal;\r\n"
				+ " text-decoration:none;\r\n" + " font-family:Arial;\r\n" + " mso-generic-font-family:auto;\r\n"
				+ " mso-font-charset:0;\r\n" + " mso-number-format:General;\r\n" + " text-align:left;\r\n"
				+ " vertical-align:bottom;\r\n" + " border-top:.5pt solid #CAC9D9;\r\n"
				+ " border-right:.5pt solid #EBEBEB;\r\n" + " border-bottom:.5pt solid #EBEBEB;\r\n"
				+ " border-left:.5pt solid #EBEBEB;\r\n" + " background:white;\r\n" + " mso-pattern:white none;\r\n"
				+ " white-space:nowrap;}\r\n" + ".xl671886\r\n" + " {padding-top:1px;\r\n" + " padding-right:1px;\r\n"
				+ " padding-left:1px;\r\n" + " mso-ignore:padding;\r\n" + " color:#333333;\r\n"
				+ " font-size:9.0pt;\r\n" + " font-weight:400;\r\n" + " font-style:normal;\r\n"
				+ " text-decoration:none;\r\n" + " font-family:Arial;\r\n" + " mso-generic-font-family:auto;\r\n"
				+ " mso-font-charset:0;\r\n" + " mso-number-format:\"\\@\";\r\n" + " text-align:left;\r\n"
				+ " vertical-align:bottom;\r\n" + " border:.5pt solid #EBEBEB;\r\n" + " background:#F8FBFC;\r\n"
				+ " mso-pattern:white none;\r\n" + " white-space:nowrap;}\r\n" + ".xl681886\r\n"
				+ " {padding-top:1px;\r\n" + " padding-right:1px;\r\n" + " padding-left:1px;\r\n"
				+ " mso-ignore:padding;\r\n" + " color:white;\r\n" + " font-size:9.0pt;\r\n" + " font-weight:700;\r\n"
				+ " font-style:normal;\r\n" + " text-decoration:none;\r\n" + " font-family:Arial;\r\n"
				+ " mso-generic-font-family:auto;\r\n" + " mso-font-charset:0;\r\n" + " mso-number-format:\"\\@\";\r\n"
				+ " text-align:left;\r\n" + " vertical-align:bottom;\r\n" + " border-top:.5pt solid #3877A6;\r\n"
				+ " border-right:.5pt solid #3877A6;\r\n" + " border-bottom:.5pt solid #A5A5B1;\r\n"
				+ " border-left:.5pt solid #3877A6;\r\n" + " background:#0B64A0;\r\n" + " mso-pattern:white none;\r\n"
				+ " white-space:normal;}\r\n" + "-->\r\n" + "</style>\r\n" + "</head>\r\n" + "\r\n" + "<body>\r\n"
				+ "<!--[if !excel]>&nbsp;&nbsp;<![endif]-->\r\n"
				+ "<!--The following information was generated by Microsoft Excel's Publish as Web\r\n"
				+ "Page wizard.-->\r\n"
				+ "<!--If the same item is republished from Excel, all information between the DIV\r\n"
				+ "tags will be replaced.-->\r\n" + "<!----------------------------->\r\n"
				+ "<!--START OF OUTPUT FROM EXCEL PUBLISH AS WEB PAGE WIZARD -->\r\n"
				+ "<!----------------------------->\r\n" + "\r\n"
				+ "<div id=\"Temp_Report2\" align=center x:publishsource=\"Excel\">\r\n" + "\r\n"
				+ "<table border=1 cellpadding=0 cellspacing=0 width=1225 style='border-collapse:\r\n"
				+ " collapse;table-layout:fixed;width:917pt'>\r\n" + "\r\n"
				+ " <tr height=48 style='height:36.0pt'>\r\n"
				+ " <td height=48 class=xl681886 width=75 style='height:36.0pt;width:56pt;text-align:center'>Product\r\n"
				+ "</td> \r\n" + "\r\n"
				+ " <td class=xl681886 width=75 style='border-left:none;width:56pt;text-align:center'>Issue Key</td>\r\n"
				+ " <td class=xl681886 width=79 style='border-left:none;width:135pt;text-align:center'>Summary</td>\r\n"
				+ " <td class=xl681886 width=75 style='border-left:none;width:56pt;text-align:center'>Status</td>\r\n"
				+ " <td class=xl681886 width=79 style='border-left:none;width:59pt;text-align:center'>Customer Name</td>\r\n"
				+ "<td class=xl681886 width=151 style='border-left:none;width:56pt;text-align:center'>CFNS Update Date</td>"
				+ "<td class=xl681886 width=151 style='border-left:none;width:56pt;text-align:center'>CFNS Update</td>"
				+ " <td class=xl681886 width=75 style='border-left:none;width:40pt;text-align:center'>Severity</td>\r\n"
				+ " <td class=xl681886 width=69 style='border-left:none;width:40pt;text-align:center'>Priority</td>\r\n"
				+ " <td class=xl681886 width=75 style='border-left:none;width:70pt;text-align:center'>Created</td>\r\n"
				+ " <td class=xl681886 width=75 style='border-left:none;width:56pt;text-align:center'>Resolved Date</td>\r\n"
				+ " <td class=xl681886 width=75 style='border-left:none;width:56pt;text-align:center'>SLA Measure</td>\r\n"
				+ " <td class=xl681886 width=151 style='border-left:none;width:56pt;text-align:center'>SLA Days Open vs Target</td>\r\n"
				+ " <td class=xl681886 width=151 style='border-left:none;width:135pt;text-align:center'>SLA Initial Response (Days in Triage) vs Target</td>\r\n"
				+ " <td class=xl681886 width=151 style='border-left:none;width:56pt;text-align:center'>Update Freq Days Since Last Update</td>\r\n";
		if(extraColumn == 1) {
			mainBody=mainBody+ " <td class=xl681886 width=151 style='border-left:none;width:56pt;text-align:center'>Assignee</td>\r\n";
		}
		mainBody=mainBody+ " </tr>";

		Map<String, String> colors = new HashMap<String, String>();
		colors.put("FF99CC00", "yellowgreen");
		colors.put("FFC0C0C0", "silver");
		colors.put("FFFF0000", "red");
		colors.put("FFFFFF00", "yellow");
		for (int i = 0; i < qlidInfo.size(); i++) {
			Map<String, String> mapOfqlidInfo = qlidInfo.get(i);
			String Product = mapOfqlidInfo.get("Product");
			String Issue_Key = mapOfqlidInfo.get("Issue Key");
			String Summary = mapOfqlidInfo.get("Summary");
			String Status = mapOfqlidInfo.get("Status");
			// String Synopsis = m.get("Synopsis");
			String Customer_Name = mapOfqlidInfo.get("Customer Name");
			String Severity = mapOfqlidInfo.get("Severity");
			String Priority = mapOfqlidInfo.get("Priority");
			String Created = mapOfqlidInfo.get("Created");
			String Resolved_Date = mapOfqlidInfo.get("Resolved Date");
			String sla_measure = mapOfqlidInfo.get("Sla Measure");
			String open_target = mapOfqlidInfo.get("SLA Days Open vs Target");
			String initialResponse = mapOfqlidInfo.get("SLA Initial Response (Days in Triage) vs Target");
			String freqdays = mapOfqlidInfo.get("Update Freq Days Since Last Update");
			String cnsUpdateDate = mapOfqlidInfo.get("CFNS Update Date");
			String cfnsUpdate = mapOfqlidInfo.get("CFNS Update");
			String assignee = mapOfqlidInfo.get("Assignee");
			String jiraLink = "https://jira.ncr.com/browse/" + Issue_Key;
			String rows = "<tr height=48 style='height:36.0pt'><td height=48 width=75 style='height:36.0pt;width:56pt;text-align:center'>"
					+ Product + "</td>" + "<td width=75 style='border-left:none;width:56pt;text-align:center'><a href="
					+ jiraLink + ">" + Issue_Key + "</a>" + "</td>"
					+ "<td width=79 style='border-left:none;width:135pt;text-align:center'>" + Summary + "</td>"
					+ "<td width=75 style='border-left:none;width:56pt;text-align:center'>" + Status + "</td>"
					+ "<td width=79 style='border-left:none;width:59pt;text-align:center'>" + Customer_Name + "</td>"
					+ "<td width=79 style='border-left:none;width:135pt;text-align:center'>" + cnsUpdateDate + "</td>"
					+ "<td width=79 style='border-left:none;width:135pt;text-align:center'>" + cfnsUpdate + "</td>"
					+ "<td width=75 style='border-left:none;width:40pt;text-align:center'>" + Severity + "</td>"
					+ "<td width=69 style='border-left:none;width:40pt;text-align:center'>" + Priority + "</td>"
					+ "<td width=75 style='border-left:none;width:70pt;text-align:center'>" + Created + "</td>"
					+ "<td width=75 style='border-left:none;width:56pt;text-align:center'>" + Resolved_Date + "</td>"
					+ "<td width=75 style='border-left:none;width:56pt;text-align:center'>" + sla_measure + "</td>";
			String colorName = mapOfqlidInfo.get((colorIndexes.get(0).toString())).toString();
			String textColor = "black";
			if (colorName.contentEquals("FFFF0000"))
				textColor = "white";
			rows = rows + "<td width=151 style='border-left:none;width:56pt;text-align:center;background-color:"
					+ colors.get(colorName) + "'" + ">" + "<p style=color:'" + textColor + "'>" + open_target
					+ "</p></td>";
			colorName = mapOfqlidInfo.get((colorIndexes.get(1).toString())).toString();
			textColor = "black";
			if (colorName.contentEquals("FFFF0000"))
				textColor = "white";
			rows = rows + "<td width=151 style='border-left:none;width:135pt;text-align:center;background-color:"
					+ colors.get(colorName) + "'" + ">" + "<p style=color:'" + textColor + "'>" + initialResponse
					+ "</p></td>";
			colorName = mapOfqlidInfo.get((colorIndexes.get(2).toString())).toString();
			textColor = "black";
			if (colorName.contentEquals("FFFF0000"))
				textColor = "white";
			rows = rows + "<td width=151 style='border-left:none;width:56pt;text-align:center;background-color:"
					+ colors.get(colorName) + "'" + ">" + "<p style=color:'" + textColor + "'>" + freqdays
					+ "</p></td>";
			if(extraColumn == 1) {
			rows = rows +"<td width=75 style='border-left:none;width:56pt;text-align:center'>" + assignee + "</td>";
			}
			rows = rows +"</tr>";
			mainBody = mainBody + rows;
		}
		mainBody = mainBody + "</table></div></body></html>";
		return mainBody;
	}

	// This function contains legend which tells about description of color codes

	private String getLegend() {

		String legend = "<html xmlns:o=\"urn:schemas-microsoft-com:office:office\"\r\n"

				+ "xmlns:x=\"urn:schemas-microsoft-com:office:excel\"\r\n"

				+ "xmlns=\"http://www.w3.org/TR/REC-html40\">\r\n" + "\r\n" + "<head>\r\n"

				+ "<meta http-equiv=Content-Type content=\"text/html; charset=windows-1252\">\r\n"

				+ "<meta name=ProgId content=Excel.Sheet>\r\n"

				+ "<meta name=Generator content=\"Microsoft Excel 15\">\r\n"

				+ "<link rel=File-List href=\"OpenSummary_files/filelist.xml\">\r\n"

				+ "<style id=\"Temp_Report2_Styles\">\r\n" + "<!--table\r\n"

				+ "             {mso-displayed-decimal-separator:\"\\.\";\r\n"

				+ "             mso-displayed-thousand-separator:\"\\,\";}\r\n" + ".xl151886\r\n"

				+ "             {padding-top:1px;\r\n" + "             padding-right:1px;\r\n"

				+ "             padding-left:1px;\r\n" + "             mso-ignore:padding;\r\n"

				+ "             color:black;\r\n" + "             font-size:10.0pt;\r\n"

				+ "             font-weight:400;\r\n" + "             font-style:normal;\r\n"

				+ "             text-decoration:none;\r\n" + "             font-family:Arial;\r\n"

				+ "             mso-generic-font-family:auto;\r\n" + "             mso-font-charset:0;\r\n"

				+ "             mso-number-format:General;\r\n" + "             text-align:general;\r\n"

				+ "             vertical-align:bottom;\r\n" + "             mso-background-source:auto;\r\n"

				+ "             mso-pattern:auto;\r\n" + "             white-space:nowrap;}\r\n" + ".xl631886\r\n"

				+ "             {padding-top:1px;\r\n" + "             padding-right:1px;\r\n"

				+ "             padding-left:1px;\r\n" + "             mso-ignore:padding;\r\n"

				+ "             color:#333333;\r\n" + "             font-size:9.0pt;\r\n"

				+ "             font-weight:400;\r\n" + "             font-style:normal;\r\n"

				+ "             text-decoration:none;\r\n" + "             font-family:Arial;\r\n"

				+ "             mso-generic-font-family:auto;\r\n" + "             mso-font-charset:0;\r\n"

				+ "             mso-number-format:\"\\@\";\r\n" + "             text-align:center;\r\n"

				+ "             vertical-align:bottom;\r\n" + "             border:.5pt solid #EBEBEB;\r\n"

				+ "             background:#F8FBFC;\r\n" + "             mso-pattern:white none;\r\n"

				+ "             white-space:nowrap;}\r\n" + ".xl641886\r\n" + "             {padding-top:1px;\r\n"

				+ "             padding-right:1px;\r\n" + "             padding-left:1px;\r\n"

				+ "             mso-ignore:padding;\r\n" + "             color:#333333;\r\n"

				+ "             font-size:9.0pt;\r\n" + "             font-weight:400;\r\n"

				+ "             font-style:normal;\r\n" + "             text-decoration:none;\r\n"

				+ "             font-family:Arial;\r\n" + "             mso-generic-font-family:auto;\r\n"

				+ "             mso-font-charset:0;\r\n" + "             mso-number-format:General;\r\n"

				+ "             text-align:right;\r\n" + "             vertical-align:bottom;\r\n"

				+ "             border:.5pt solid #EBEBEB;\r\n" + "             background:#F8FBFC;\r\n"

				+ "             mso-pattern:white none;\r\n" + "             white-space:nowrap;}\r\n" + ".xl651886\r\n"

				+ "             {padding-top:1px;\r\n" + "             padding-right:1px;\r\n"

				+ "             padding-left:1px;\r\n" + "             mso-ignore:padding;\r\n"

				+ "             color:#333333;\r\n" + "             font-size:9.0pt;\r\n"

				+ "             font-weight:700;\r\n" + "             font-style:normal;\r\n"

				+ "             text-decoration:none;\r\n" + "             font-family:Arial;\r\n"

				+ "             mso-generic-font-family:auto;\r\n" + "             mso-font-charset:0;\r\n"

				+ "             mso-number-format:General;\r\n" + "             text-align:right;\r\n"

				+ "             vertical-align:bottom;\r\n" + "             border-top:.5pt solid #CAC9D9;\r\n"

				+ "             border-right:.5pt solid #EBEBEB;\r\n"

				+ "             border-bottom:.5pt solid #EBEBEB;\r\n"

				+ "             border-left:.5pt solid #EBEBEB;\r\n" + "             background:white;\r\n"

				+ "             mso-pattern:white none;\r\n" + "             white-space:nowrap;}\r\n" + ".xl661886\r\n"

				+ "             {padding-top:1px;\r\n" + "             padding-right:1px;\r\n"

				+ "             padding-left:1px;\r\n" + "             mso-ignore:padding;\r\n"

				+ "             color:#333333;\r\n" + "             font-size:9.0pt;\r\n"

				+ "             font-weight:700;\r\n" + "             font-style:normal;\r\n"

				+ "             text-decoration:none;\r\n" + "             font-family:Arial;\r\n"

				+ "             mso-generic-font-family:auto;\r\n" + "             mso-font-charset:0;\r\n"

				+ "             mso-number-format:General;\r\n" + "             text-align:left;\r\n"

				+ "             vertical-align:bottom;\r\n" + "             border-top:.5pt solid #CAC9D9;\r\n"

				+ "             border-right:.5pt solid #EBEBEB;\r\n"

				+ "             border-bottom:.5pt solid #EBEBEB;\r\n"

				+ "             border-left:.5pt solid #EBEBEB;\r\n" + "             background:white;\r\n"

				+ "             mso-pattern:white none;\r\n" + "             white-space:nowrap;}\r\n" + ".xl671886\r\n"

				+ "             {padding-top:1px;\r\n" + "             padding-right:1px;\r\n"

				+ "             padding-left:1px;\r\n" + "             mso-ignore:padding;\r\n"

				+ "             color:#333333;\r\n" + "             font-size:9.0pt;\r\n"

				+ "             font-weight:400;\r\n" + "             font-style:normal;\r\n"

				+ "             text-decoration:none;\r\n" + "             font-family:Arial;\r\n"

				+ "             mso-generic-font-family:auto;\r\n" + "             mso-font-charset:0;\r\n"

				+ "             mso-number-format:\"\\@\";\r\n" + "             text-align:left;\r\n"

				+ "             vertical-align:bottom;\r\n" + "             border:.5pt solid #EBEBEB;\r\n"

				+ "             background:#F8FBFC;\r\n" + "             mso-pattern:white none;\r\n"

				+ "             white-space:nowrap;}\r\n" + ".xl681886\r\n" + "             {padding-top:1px;\r\n"

				+ "             padding-right:1px;\r\n" + "             padding-left:1px;\r\n"

				+ "             mso-ignore:padding;\r\n" + "             color:white;\r\n"

				+ "             font-size:9.0pt;\r\n" + "             font-weight:700;\r\n"

				+ "             font-style:normal;\r\n" + "             text-decoration:none;\r\n"

				+ "             font-family:Arial;\r\n" + "             mso-generic-font-family:auto;\r\n"

				+ "             mso-font-charset:0;\r\n" + "             mso-number-format:\"\\@\";\r\n"

				+ "             text-align:left;\r\n" + "             vertical-align:bottom;\r\n"

				+ "             border-top:.5pt solid #3877A6;\r\n"

				+ "             border-right:.5pt solid #3877A6;\r\n"

				+ "             border-bottom:.5pt solid #A5A5B1;\r\n"

				+ "             border-left:.5pt solid #3877A6;\r\n" + "             background:#0B64A0;\r\n"

				+ "             mso-pattern:white none;\r\n" + "             white-space:normal;}\r\n" + "-->\r\n"

				+ "</style>\r\n" + "</head>\r\n" + "\r\n" + "<body>\r\n"

				+ "<!--[if !excel]>&nbsp;&nbsp;<![endif]-->\r\n"

				+ "<!--The following information was generated by Microsoft Excel's Publish as Web\r\n"

				+ "Page wizard.-->\r\n"

				+ "<!--If the same item is republished from Excel, all information between the DIV\r\n"

				+ "tags will be replaced.-->\r\n" + "<!----------------------------->\r\n"

				+ "<!--START OF OUTPUT FROM EXCEL PUBLISH AS WEB PAGE WIZARD -->\r\n"

				+ "<!----------------------------->\r\n"

				+ "<div id=\"Temp_Report2\" align=center x:publishsource=\"Excel\">\r\n" + "\r\n"

				+ "<table border=1 cellpadding=0 cellspacing=0  style='border-collapse:\r\n"

				+ " collapse;table-layout:fixed;width:500pt'>\r\n" + " <tr height=48 >\r\n"

				+ " <td class=xl681886 width=75 style='width:56pt;align:center;text-align:center'>Legend</td>\r\n"

				+ "  <td class=xl681886 width=75 style='width:56pt;align:center;text-align:center'>SLA Days Open vs Target</td>\r\n"

				+ "  <td class=xl681886 width=75 style='width:56pt;align:center;text-align:center'>SLA Initial Response (Days in Triage) vs Target</td>\r\n"

				+ "  <td class=xl681886 width=75 style='width:56pt;align:center;text-align:center'>Update Freq Days Since Last Update</td>\r\n"

				+ "  \r\n" + "  \r\n" + "  </tr>\r\n" + "  <tr height=48 >\r\n"

				+ " <td  width=75 style='width:56pt;align:center;text-align:center'>Action Required</td>\r\n"

				+ "  <td  width=75 style='width:56pt;align:center;text-align:center'>Provide Fix to the customer and move Jira Status based on \"SLA Measure\"</td>\r\n"

				+ "  <td  width=75 style='width:56pt;align:center;text-align:center'>Provide Initial Response  for the Jira item moving it from Not Started -> In Triage -> Next Status</td>\r\n"

				+ "  <td  width=75 style='width:56pt;align:center;text-align:center'>Update \"Customer Facing Next Steps\" field on a periodic basis as defined by the SLA target </td>\r\n"

				+ "  \r\n" + "  \r\n" + "  </tr>\r\n" + "  \r\n" + "  <tr height=48 >\r\n"

				+ " <td  width=75 style='width:56pt;align:center;text-align:center'>Green</td>\r\n"

				+ "  <td  width=75 colspan=\"3\" style='width:56pt;align:center;background-color:yellowgreen;text-align:center'>Still within target more than 2 days</td>\r\n"

				+ " \r\n" + "  \r\n" + "  </tr>\r\n" + "  <tr height=48 >\r\n"

				+ " <td  width=75 style='width:56pt;align:center;text-align:center'>Yellow</td>\r\n"

				+ "  <td  width=75 colspan=\"3\" style='width:56pt;align:center;background-color:yellow;text-align:center'>Action Required now within 2 days of target</td>\r\n"

				+ "\r\n" + "  </tr>\r\n" + "  <tr height=48 >\r\n"

				+ " <td  width=75 style='width:56pt;align:center;text-align:center'>Red</td>\r\n"

				+ "  <td  width=75 colspan=\"3\" style='width:56pt;align:center;background-color:red;text-align:center;text-color:green'><p style='color:white'>Target Missed</td>\r\n"

				+ " \r\n" + "  </tr>\r\n" + "  \r\n" + "   <tr height=48 >\r\n"

				+ " <td  width=75 style='width:56pt;align:center;text-align:center'>Grey</td>\r\n"

				+ "  <td  width=75 style='width:56pt;align:center'>&nbsp</td>\r\n"

				+ "  <td  width=75 style='width:56pt;align:center;background-color:silver;text-align:center'>Initial Response Provided</td>\r\n"

				+ "  <td  width=75 style='width:56pt;align:center'>&nbsp </td>\r\n" + "  \r\n" + "  \r\n"

				+ "  </tr>\r\n" + "  \r\n" + "  </table>\r\n" + "\r\n" + "</div>\r\n" + "\r\n" + "\r\n"

				+ "<!----------------------------->\r\n"

				+ "<!--END OF OUTPUT FROM EXCEL PUBLISH AS WEB PAGE WIZARD-->\r\n"

				+ "<!----------------------------->\r\n" + "</body>\r\n" + "\r\n" + "</html>\r\n" + "\r\n" + "";

		return legend;

	}

	// In this function we are constructing body for each assignes's defects

	// Not using this function

	public static Map<String, Integer> getColumnNumbers(Sheet sheet, int[] rowNumbers, List<String> columnNames) {

		// A Map of column names and numbers. Column Name forms the key whereas the

		// column number forms the value.

		Map<String, Integer> columnNumbers = new HashMap<String, Integer>();

		Row headerRow1 = sheet.getRow(rowNumbers[0]);

		Row headerRow2 = sheet.getRow(rowNumbers[1]);

		// Fetch the number of merged regions (i.e., column groupings)

		int mergedRegionCount = sheet.getNumMergedRegions();

		// Build a map of starting column number and the number of columns for each

		// column grouping

		Map<Integer, Integer> rangeMap = new HashMap<Integer, Integer>();

		for (int count = 0; count < mergedRegionCount; count++) {

			CellRangeAddress range = sheet.getMergedRegion(count);

			rangeMap.put(Integer.valueOf(range.getFirstColumn()), Integer.valueOf(range.getNumberOfCells()));

		}

		// Get the minimum and maximum column indexes of the row

		short minColIndex = headerRow2.getFirstCellNum();

		short maxColIndex = headerRow2.getLastCellNum();

		String colGroupName = "";

		int headerColCount = 0;

		// Iterate over the column index range

		for (int colIndex = minColIndex; colIndex < maxColIndex; colIndex++) {

			if (rangeMap.containsKey(Integer.valueOf(colIndex))) {

				Cell headerCell = headerRow1.getCell(colIndex);

				colGroupName = getCellValue(headerCell);

				headerColCount = rangeMap.get(Integer.valueOf(colIndex));

			}

			// Fetch the cell for the specified column index

			Cell cell = headerRow2.getCell(colIndex);

			// If the cell is null, we ignore it.

			if (cell == null) {

				continue;

			}

			// Get the cell value which is actually the column name

			String columnName = getCellValue(cell);

			if (headerColCount-- > 0) {

				columnName = colGroupName + "." + columnName;

			}

			// If the column name is in the list of column names then we record the column

			// index for that column name

			if (columnNames.contains(columnName)) {

				columnNumbers.put(columnName, Integer.valueOf(colIndex));

			}

		}

		return columnNumbers;

	}

	// This returns cloumn names mapped by their position(i.e column number) in

	// excel file

}
