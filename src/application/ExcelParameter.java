package application;

import java.io.File;

public class ExcelParameter {

	public static int columns = 3;
	public static int sheet = 0;
	public static int sheets = 78;
	public static String excelFileNameIn = "Nouns";
	public static String excelFileSheetNameOut = "Nouns";
	public static String excelFilePathOut = System.getProperty("user.home") + File.separator + "Desktop" + File.separator + excelFileSheetNameOut + "_Excersice_List.xlsx";
	public static String excelFilePathOutI = System.getProperty("user.home") + File.separator + "Desktop" + File.separator + excelFileSheetNameOut + "_Excersice_I.xlsx";
	public static String excelFilePathOutII = System.getProperty("user.home") + File.separator + "Desktop" + File.separator + excelFileSheetNameOut + "_Excersice_II.xlsx";
	public static String excelFilePathIn = ".\\src\\" + ExcelParameter.excelFileNameIn + ".xlsx";
	public static String excelSheetNameOne = "List";
	public static String excelSheetNameTwo = "Basic";
	public static String excelSheetNameThree = "Mixed";

	public static String fontNameExcel = "Segoe UI";
	public static String sheetName = "Exercise";

}
