package xsl2vCard.maker;

import java.io.BufferedWriter;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.FileWriter;
import java.io.IOException;

import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.*;


/**
 * Hello world!
 *
 */
public class App {
	public static int TOTALROWS = 250;
	public static void main(String[] args) throws InvalidFormatException, IOException {



		String path = "/path/to/your/xsl/file/with/contacts.xls";

		File myFile = new File(path);
		String parentFolder = "newVCards";
		System.out.println(myFile.exists());
		FileInputStream fis = new FileInputStream(myFile);
		HSSFWorkbook myWorkBook = new HSSFWorkbook(fis);
		HSSFSheet mySheet = myWorkBook.getSheetAt(0);

		Row row;
		Cell cell;
		String fname;
		String lname;
		String number1;
		String number2;
		String number3;

		for (int i = 0; i < TOTALROWS; i ++) {
			row = mySheet.getRow(i);

			try {
				fname = row.getCell(0).getStringCellValue();}
			catch(Exception e) {
				fname = " ";}
			try {
				lname = row.getCell(1).getStringCellValue();}
			catch(Exception e) {
				lname = " ";
			}
			try {
				number1 = row.getCell(2).getStringCellValue();}
			catch(Exception e) {
				number1 = " ";
			}
			try {
				number2 = row.getCell(3).getStringCellValue();}
			catch(Exception e) {
				number2 = " ";
			}
			try {
				number3 = row.getCell(4).getStringCellValue();}
			catch(Exception e) {
				number3 = " ";
			}

			System.out.print(fname+" "+lname+" "+number1+" "+number2+" "+number3+"\n");
			createVCardForGivenName(fname,lname, number1, number2, number3);
		}

		myWorkBook.close();
	}

	public static void createVCardForGivenName(String fname, String lname, String number1, String number2, String number3) throws IOException {

		File newVCard = new File("/path/to/save/vCard/files/"+fname+lname+".vCard");
		StringBuilder sb= new StringBuilder();
		sb.append("\nBEGIN:VCARD");
		sb.append("\nVERSION:4.0");
		sb.append("\nN:"+fname+";");
		sb.append(lname+";");
		sb.append("\nFN:"+fname+" "+lname);
		sb.append("\nTEL;TYPE=home,voice;VALUE=uri:"+number1+"");
		//sb.append("\nTEL;TYPE=home,voice;VALUE=uri:"+number2+"");
		//sb.append("\nTEL;TYPE=home,voice;VALUE=uri:"+number3+"");
		sb.append("\nREV:20080424T195243Z");
		sb.append("\nx-qq:21588891");
		sb.append("\nEND:VCARD");

		BufferedWriter writer = null;

		writer = new BufferedWriter( new FileWriter(newVCard));
		writer.write(sb.toString());

		writer.close();
	}



}
