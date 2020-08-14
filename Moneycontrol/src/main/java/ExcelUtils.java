 
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Date;
import java.util.HashMap;
import java.util.Map;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import io.restassured.RestAssured;
import io.restassured.response.Response;
import io.restassured.specification.RequestSpecification;

public class ExcelUtils {
	static ArrayList<String> stockcodes = new ArrayList<String>();
	static Map<String, Object[]> data = new HashMap<String, Object[]>();
	//static String excelFilePath = "C:\\Users\\Akshay\\Desktop\\stocks.xlsx";
	static String excelFilePath = "C:\\Users\\Dwaghmod\\OneDrive - Capgemini\\Desktop\\stoock.xlsx";


	final int sheetNo= 2;
	public static void main(String[] args) throws Exception
	{
		System.out.println(excelFilePath+"**********");
			ExcelUtils excel = new ExcelUtils();
			stockcodes = excel.readExcel();
			data = excel.getData(stockcodes);
			excel.writeExcel(data);

	} 

	public ArrayList<String> readExcel() throws Exception {
		try {
			FileInputStream inputStream = new FileInputStream(new File(excelFilePath));
			Workbook workbook = new XSSFWorkbook(inputStream);
			XSSFSheet sheet = (XSSFSheet) workbook.getSheetAt(sheetNo);

			for (int rowIndex = 0; rowIndex <= sheet.getLastRowNum(); rowIndex++) {
				Row row = sheet.getRow(rowIndex);
				Cell cell = row.getCell(1);
				System.out.println(cell.getStringCellValue());
				stockcodes.add(cell.getStringCellValue());
			}

			return stockcodes;
		} catch (Exception e) {
			throw new Exception("error in reading "+e.getMessage());
		}



	}

	public Map<String, Object[]> getData(ArrayList<String> stockcodes) throws Exception {
		for (int i = 1; i < stockcodes.size(); i++) {
			RestAssured.baseURI = "https://priceapi.moneycontrol.com/pricefeed/nse/equitycash/";
			RequestSpecification httpRequest = RestAssured.given();
			Response response = httpRequest.get(stockcodes.get(i).toString());
			int statusCode = response.getStatusCode();
			Map<String, String> company = response.jsonPath().getMap("data");
			System.out.println(company.get("company") + "," + company.get("symbol") + "," + company.get("pricecurrent")
					+ "," + company.get("52H")+ "," + company.get("52L")+","+company.get("SC_SUBSEC")+company.get("HP")+company.get("LP"));

			data.put(stockcodes.get(i).toString(), new Object[] { company.get("company"), company.get("symbol"),
					company.get("pricecurrent"),company.get("SC_SUBSEC"),company.get("HP"),company.get("LP")});

		}
		return data;

	}

	public void writeExcel(Map<String, Object[]> data) throws IOException {
		FileInputStream inputStream = new FileInputStream(new File(excelFilePath));
		Workbook workbook = new XSSFWorkbook(inputStream);
		XSSFSheet spreadsheet = (XSSFSheet) workbook.getSheetAt(sheetNo);
		XSSFRow row;

		for (int rowIndex = 1; rowIndex <= spreadsheet.getLastRowNum(); rowIndex++) {
			row = spreadsheet.getRow(rowIndex);
			Cell cell = row.getCell(1);
			System.out.println(cell.getStringCellValue()+"                     ..........");
			Object[] objectArr = data.get(cell.getStringCellValue());
			/*for(int i=0;i<objectArr.length;i++)
			{
				System.out.println(objectArr[i]);
			}*/
			int cellid = 4;

			for (int i=0;i<objectArr.length;i++)
			{


				System.out.println();
				if(i==2||i==3||i==4||i==5||i==6||i==7)
				{

				Cell cell1 = row.createCell(cellid);
				Object obj = objectArr[i].toString();
				if (obj instanceof Date)
		            cell1.setCellValue((Date) obj);
		        else if (obj instanceof Boolean)
		            cell1.setCellValue((Boolean) obj);
		        else if (obj instanceof String)
		            cell1.setCellValue(obj.toString());
		        else if (obj instanceof Double)
		            cell1.setCellValue((Double) obj);
				cellid++;
            }
			}

			}

		FileOutputStream out = new FileOutputStream(new File(excelFilePath));

		workbook.write(out);
		out.close();
		System.out.println("Writesheet.xlsx written successfully");
	}
}
