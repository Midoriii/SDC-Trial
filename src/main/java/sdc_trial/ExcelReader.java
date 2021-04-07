package sdc_trial;

import java.io.File;
import java.io.IOException;

import org.apache.commons.math3.primes.Primes;
import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;

/**
 * 
 * The purpose of this class is to find and print out prime numbers
 * found in a provided Excel file (with some constraints). I chose
 * to implement the methods as static, as I don't see the value in
 * instantiating this class for this task.
 * 
 * @author Štěpán Beneš
 *
 */
public final class ExcelReader {
	
	private ExcelReader() {}
	
	/**
	 * The main workhorse of this class. Given a path to a valid Excel file,
	 * it opens the file and searches for prime numbers in the column B.
	 * Implicitly only the first worksheet is searched. Although extending this method
	 * to browse several sheets and possibly even columns would be trivial.
	 * 
	 * @param path - String containing the path to a .xlsx file, unprotected by a password.
	 */
	public static void findPrimes(String path) {
		Workbook workbook = null;
		
		try {
			File file = new File(path);
			workbook = WorkbookFactory.create(file);
			Sheet sheet = workbook.getSheetAt(0);
			parseSheet(sheet);
		}
		catch (IOException e) {
			printError("Given file could not be read.");
		}
		catch (EncryptedDocumentException e) {
			printError("Given file is password protected.");
		}
		/* 
		 * Either NullPointer or IllegalArgument, neither of which should happen,
		 * as the String path does exist when given to File constructor, and the workbook
		 * should have at least the first sheet.
		 */
		catch (Exception e) {
		    e.printStackTrace();
		}
		finally {
			if (workbook != null) {
				try {
					workbook.close();
				} catch (IOException e) {
					printError("Failed to close the file properly.");
				}
			}
		}
	}
	
	/**
	 * Method for cycling through the rows of a given Sheet. For each row
	 * the cell in column B is forwarded to the method parseCell(Cell c).
	 * 
	 * @param sheet - Excel Sheet with interesting data in column B.
	 */
	private static void parseSheet(Sheet sheet) {
		for (Row row : sheet) {
			Cell cell = row.getCell(Column.B.getValue());
			
			if(cell != null) {
				parseCell(cell);
			}
		}
	}
	
	/**
	 * Method for handling a single Cell. Only String type cells concern us,
	 * but the method can easily be extended to handle other types as well.
	 * Prints the cell's value to the stdout if it's a prime number.
	 * 
	 * Checking whether the value in the given cell is a prime number is done
	 * using Apache Commons Math library, as I find that more appealing than reinventing
	 * the traditional approach with square root of the value and modulo checking.
	 * 
	 * @param cell - Excel Cell of String type.
	 */
	private static void parseCell(Cell cell) {
		if(cell.getCellType().equals(CellType.STRING)) {
			String value = cell.getStringCellValue();
			
			// Taking advantage of short-circuit evaluation of &&
			if(isNumber(value) && Primes.isPrime(Integer.parseInt(value))) {
				System.out.println(value);
			}
		}
	}
	
	/**
	 * Helper for checking whether the given String is a viable Integer.
	 * Fairly simple but in my opinion prettier than try - catch used with
	 * parseInt method.
	 * 
	 * @param value - String to be checked whether it can be parsed as Integer. 
	 * @return True/False depending on the outcome of the check.
	 */
	private static boolean isNumber(String value) {
		boolean isNumber = true;
		
		for (char c : value.toCharArray()) {
			isNumber = isNumber && Character.isDigit(c);
		}
		
		return isNumber;
	}
	
	/**
	 * Helper function to consume exceptions - users generally don't need to see
	 * the stack trace, simple message and exit is enough.
	 * Could've used Logger instead of println, but that'd be an overkill here.
	 * 
	 * @param msg - Message describing what went wrong.
	 */
	private static void printError(String msg) {
		System.out.println(msg);
		System.exit(0);
	}

	/**
	 * The wise author of Effective Java advises to use Enums instead of
	 * Integer constants, so the same approach is taken here.
	 */
	private enum Column{
		B(1);
		
		private int value;
		
		Column(int value) { this.value = value; }
		
		public int getValue() { return this.value; }
	}
}
