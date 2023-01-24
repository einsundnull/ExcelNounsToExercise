package application;

import java.awt.Toolkit;
import java.awt.datatransfer.Clipboard;
import java.awt.datatransfer.StringSelection;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.Collections;
import java.util.HashSet;
import java.util.LinkedList;
import java.util.Random;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ExcelReader {

	static LinkedList<String[]> data = new LinkedList<>();
	static int wordCount = 0;
	static String toClipboard = "";

	public static LinkedList<String[]> readExcelSingleSheet(File file) {
		try {
			FileInputStream inputStream = new FileInputStream(new File(ExcelParameter.excelFilePathIn));
			Workbook workbook = new XSSFWorkbook(inputStream);
			try {
				Sheet sheet = workbook.getSheetAt(ExcelParameter.sheet);
				ExcelParameter.excelFileSheetNameOut = sheet.getSheetName();
				ExcelParameter.excelFilePathOut = System.getProperty("user.home") + File.separator + "Desktop" + File.separator + ExcelParameter.excelFileSheetNameOut
						+ "_Excersice_List.xlsx";
				ExcelParameter.excelFilePathOutI = System.getProperty("user.home") + File.separator + "Desktop" + File.separator + ExcelParameter.excelFileSheetNameOut
						+ "_Excersice_I.xlsx";
				ExcelParameter.excelFilePathOutII = System.getProperty("user.home") + File.separator + "Desktop" + File.separator + ExcelParameter.excelFileSheetNameOut
						+ "_Excersice_II.xlsx";

				for (Row row : sheet) {
					String[] cells = new String[ExcelParameter.columns];
					for (int i = 0; i < 3; i++) {
						Cell cell = row.getCell(i);
						try {
							cells[i] = cell.getStringCellValue();
						} catch (Exception e) {
							cells[i] = "Empty";
						}
						System.out.print(cells[i] + " ");
					}
					System.out.println("");
					data.add(cells);
					wordCount++;
				}
			} catch (Exception e) {
				e.printStackTrace();
			}
			workbook.close();
		} catch (IOException e) {
			e.printStackTrace();
		}
		System.out.println(wordCount);
//		ExcelParameter.sheet = 0;
		wordCount = 0;
		return data;
	}

	public static LinkedList<String[]> readExcelAllSheets(File file) {

		try {
			String excelFilePath = ".\\src\\Nouns.xlsx";
			FileInputStream inputStream = new FileInputStream(new File(excelFilePath));
//    		FileInputStream inputStream = new FileInputStream(file);
			Workbook workbook = new XSSFWorkbook(inputStream);
			try {
				for (int s = 0; s < ExcelParameter.sheets; s++) {
					Sheet sheet = workbook.getSheetAt(ExcelParameter.sheet);
					ExcelParameter.excelFileSheetNameOut = sheet.getSheetName();
					ExcelParameter.excelFilePathOut = System.getProperty("user.home") + File.separator + "Desktop" + File.separator + ExcelParameter.excelFileSheetNameOut
							+ "_Excersice_List.xlsx";
					ExcelParameter.excelFilePathOutI = System.getProperty("user.home") + File.separator + "Desktop" + File.separator + ExcelParameter.excelFileSheetNameOut
							+ "_Excersice_I.xlsx";
					ExcelParameter.excelFilePathOutII = System.getProperty("user.home") + File.separator + "Desktop" + File.separator + ExcelParameter.excelFileSheetNameOut
							+ "_Excersice_II.xlsx";
					for (Row row : sheet) {
						String[] cells = new String[ExcelParameter.columns];
						for (int i = 0; i < 3; i++) {
							Cell cell = row.getCell(i);
							try {
								cells[i] = cell.getStringCellValue();
							} catch (Exception e) {
								cells[i] = "Empty";
							}
							System.out.print(cells[i] + " ");
						}
						System.out.println("");
						data.add(cells);
						wordCount++;
					}
					ExcelParameter.sheet++;
				}

			} catch (Exception e) {
				e.printStackTrace();
			}
			workbook.close();
		} catch (IOException e) {
			e.printStackTrace();
		}
		System.out.println(wordCount);
		ExcelParameter.sheet = 0;
		wordCount = 0;
		return data;
	}

	public static int countDuplicateValues(LinkedList<String[]> linkList) {
		int count = 0;
		HashSet<String> set = new HashSet<>();
		for (String[] arr : linkList) {
			if (set.contains(arr[1])) {
				count++;
			} else {
				set.add(arr[1]);
			}
		}
		return count;
	}

//	public static LinkedList<String[]> processDataAsLinkedList(LinkedList<String[]> data) {
//		String words = "";
//		
//		String[] newArray = null;
//		LinkedList<String[]> dataNew = new LinkedList<>();
//		for (int i = 0; i < data.size(); i++) {
//			String word = "";
//			for (int n = 0; n < data.get(i).length; n++) {
//				try {
//					word =  data.get(i)[n];
//					if (n > 0) {
//						words = words + " " + word;
//					} else {
//						words = words +  word;
//					}
//				
//					words = words.replace("  ", " ");
//					if (n > 0) {
//						words = words.replace(" ", "\t");
//					}
//				} catch (Exception e) {
//					e.printStackTrace();
//					break;
//				}
//			}
////			words = words + "\n";
//			newArray = words.split("\\s+");
//			System.out.println("List contains: " + newArray.length + " words");
//			dataNew.add(newArray);
//		}
//		System.out.println(words);
//		return dataNew;
//	}

	public static LinkedList<String[]> processDataAsLinkedListSplittWhitespace(LinkedList<String[]> data) {

		String[] newArray = null;
		LinkedList<String[]> dataNew = new LinkedList<>();
		for (int i = 0; i < data.size(); i++) {
			String word = "";
			for (int n = 0; n < data.get(i).length; n++) {
				try {
//					word = data.get(i)[n];
					if (n > 0) {
						word = word + " " + data.get(i)[n];
					} else {
						word = word + data.get(i)[n];
					}
					word = word.replace("  ", " ");
				} catch (Exception e) {
					e.printStackTrace();
					break;
				}
			}
			newArray = word.split("\\s+");
			System.out.println("List contains: " + newArray.length + " words");
			dataNew.add(newArray);
		}
		return dataNew;
	}

	public static String processDataToSingleString(LinkedList<String[]> data) {
		String words = "";
		for (int i = 0; i < data.size(); i++) {
			for (int n = 0; n < data.get(i).length; n++) {
				try {
					if (n > 0) {
						words = words + "\t" + data.get(i)[n];
					} else {
						words = words + data.get(i)[n];
					}
				} catch (Exception e) {
					e.printStackTrace();
					break;
				}
			}
			words = words + "\n";
		}
		System.out.println(words);
		return words;
	}

	public static void wordsToClipboard(String words) {
		StringSelection stringSelection = new StringSelection(words);
		Clipboard clipboard = Toolkit.getDefaultToolkit().getSystemClipboard();
		clipboard.setContents(stringSelection, null);
	}

	public static LinkedList<String[]> processDataToExerciseOne(LinkedList<String[]> data) {
		LinkedList<String[]> result = new LinkedList<String[]>();

		for (int i = 0; i < data.size(); i++) {
			try {
				String[] p_0 = data.get(i);
				String[] p_1 = new String[9];
				p_1[0] = p_0[2];
				p_1[1] = replaceByDots(p_0[1]);
				p_1[2] = p_0[1];
				p_1[3] = "";
				p_1[4] = p_0[3];
				p_1[5] = replaceByDots(p_0[4]);

				p_1[6] = p_0[4];
				p_1[7] = "";
				p_1[8] = p_0[0];
				result.add(p_1);
			} catch (Exception e) {
				e.printStackTrace();
			}
		}
		return result;
	}

	public static String createString(LinkedList<String[]> data) {
		String words = "";
		for (int i = 0; i < data.size(); i++) {
			for (int n = 0; n < data.size(); n++) {
				words = words + data.get(i)[n] + "\t";
			}
			words = words + "\n";
		}
		System.out.println(words);
		return words;

	}

	public static String replaceByDots(String word) {

		char[] chars = word.toCharArray();
		String replace = "";
		for (int i = 0; i < chars.length; i++) {
			replace = replace + "_";
		}

		return replace;
	}

	public static LinkedList<String[]> processDataToExerciseTwo(LinkedList<String[]> data) {
		String add = " (pl.)";
		LinkedList<String[]> dataExercise = new LinkedList<>();
		String[] sideOne = new String[3];
		String[] sideTwo = new String[3];
//	    for (String[] row : data) {
		for (int i = 0; i < data.size(); i++) {
			String[] row = data.get(i);
			String[] rowInverse = data.get((data.size() - 1) - i);
			String one = row[0];
			String oneInverse = rowInverse[0];
			String two = row[1];
			String twoInverse = rowInverse[1];
			String three = row[2];
			String threeInverse = rowInverse[2];
			String four = oneInverse + add;
			String fourInverse = oneInverse + add;

			String[] exerciseTwo_I = new String[7];
			exerciseTwo_I[0] = twoInverse;
			exerciseTwo_I[1] = replaceByDots(oneInverse);
			exerciseTwo_I[2] = one;
			exerciseTwo_I[3] = "";
			exerciseTwo_I[4] = three;
			exerciseTwo_I[5] = replaceByDots(four);
			exerciseTwo_I[6] = four;
			dataExercise.add(exerciseTwo_I);

			String[] exerciseTwo_II = new String[7];
			exerciseTwo_II[0] = one;
			exerciseTwo_II[1] = replaceByDots(two);
			exerciseTwo_II[2] = two;
			exerciseTwo_II[3] = "";
			exerciseTwo_II[4] = threeInverse;
			exerciseTwo_II[5] = replaceByDots(fourInverse);
			exerciseTwo_II[6] = fourInverse;
			dataExercise.add(exerciseTwo_II);
		}

		for (int i = 0; i < dataExercise.size(); i++) {
			boolean rnd = new Random().nextBoolean();
			if(rnd) {
				
			}
			sideOne[0] = dataExercise.get(i)[0];
			sideOne[1] = dataExercise.get(i)[1];
			sideOne[2] = dataExercise.get(i)[2];
			sideTwo[4] = dataExercise.get(i)[4];
			sideTwo[5] = dataExercise.get(i)[5];
			sideTwo[6] = dataExercise.get(i)[6];

			sideTwo[0] = dataExercise.get(i)[4];
			sideTwo[1] = dataExercise.get(i)[5];
			sideTwo[2] = dataExercise.get(i)[6];
			sideOne[4] = dataExercise.get(i)[0];
			sideOne[5] = dataExercise.get(i)[1];
			sideOne[6] = dataExercise.get(i)[2];
		}

		Collections.shuffle(dataExercise);
		return dataExercise;
	}

//	public static LinkedList<String[]> processDataToExerciseTwo(LinkedList<String[]> data) {
//		AI WORIKING
//	    String add = " (pl.)";
//	    LinkedList<String[]> dataExercise = new LinkedList<>();
//
//	    for (String[] row : data) {
//	        String one = row[0];
//	        String two = row[1];
//	        String three = row[2];
//	        String four = one + add;
//
//	        String[] exerciseTwo_I = new String[7];
//	        exerciseTwo_I[0] = two;
//	        exerciseTwo_I[1] = replaceByDots(one);
//	        exerciseTwo_I[2] = one;
//	        exerciseTwo_I[3] = "";
//	        exerciseTwo_I[4] = three;
//	        exerciseTwo_I[5] = replaceByDots(four);
//	        exerciseTwo_I[6] = four;
//	        dataExercise.add(exerciseTwo_I);
//
//	        String[] exerciseTwo_II = new String[7];
//	        exerciseTwo_II[0] = one;
//	        exerciseTwo_II[1] = replaceByDots(two);
//	        exerciseTwo_II[2] = two;
//	        exerciseTwo_II[3] = "";
//	        exerciseTwo_II[4] = three + add;
//	        exerciseTwo_II[5] = replaceByDots(four);
//	        exerciseTwo_II[6] = four;
//	        dataExercise.add(exerciseTwo_II);
//	    }
//	    
//	    Collections.shuffle(dataExercise);
//	    return dataExercise;
//	}

//	public static LinkedList<String[]> processDataToExerciseTwo(LinkedList<String[]> data) {
//		String one = "";
//		String two = "";
//		String three = "";
//		String four = "";
//		String add = " (pl.)";
//
//		LinkedList<String[]> dataExercise = new LinkedList<>();
//		for (int i = 0; i < data.size(); i++) {
//			String[] exerciseTwo_I = new String[7];
//			one = data.get(i)[0];
//			two = data.get(i)[1];
//			three = data.get(i)[2];
//			four = one + add;
//
//			exerciseTwo_I[0] = two;
//			exerciseTwo_I[1] = replaceByDots(one);
//			exerciseTwo_I[2] = one;
//			exerciseTwo_I[3] = "";
//			exerciseTwo_I[4] = three;
//			exerciseTwo_I[5] = replaceByDots(four);
//			exerciseTwo_I[6] = four;
//			dataExercise.add(exerciseTwo_I);
//		}
//		mixList(data);
//		for (int i = 0; i < data.size(); i++) {
//			String[] exerciseTwo_II = new String[7];
//			one = data.get(i)[1];
//			two = data.get(i)[0];
//			three = two + add;
//			four = data.get(i)[2];
//
//			exerciseTwo_II[0] = one;
//			exerciseTwo_II[1] = replaceByDots(two);
//			exerciseTwo_II[2] = two;
//			exerciseTwo_II[3] = "";
//			exerciseTwo_II[4] = three;
//			exerciseTwo_II[5] = replaceByDots(four);
//			exerciseTwo_II[6] = four;
//			dataExercise.add(exerciseTwo_II);
//		}
//
//		return dataExercise;
//	}

	public static LinkedList<String[]> mixList(LinkedList<String[]> list) {
		Random rand = new Random();
		for (int i = list.size() - 1; i > 0; i--) {
			int j = rand.nextInt(i);
			String[] temp = list.get(i);
			list.set(i, list.get(j));
			list.set(j, temp);
		}
		return list;
	}

//	public static XSSFWorkbook createExcelSheetWithCellWidth(String sheetName, String outputFilePath, String fontName, short fontSize) {
//		XSSFWorkbook workbook = new XSSFWorkbook();
//		Sheet sheet = workbook.createSheet(sheetName);
//
//		// Set the width of each cell in centimeters
//		sheet.setColumnWidth(0, (int) (3.75 * 37)); // 3.75cm
//		sheet.setColumnWidth(1, (int) (2.00 * 37)); // 2.00cm
//		sheet.setColumnWidth(2, (int) (0.75 * 37)); // 0.75cm
//		sheet.setColumnWidth(3, (int) (2.00 * 37)); // 2.00cm
//		sheet.setColumnWidth(4, (int) (0.75 * 37)); // 0.75cm
//		sheet.setColumnWidth(5, (int) (3.75 * 37)); // 3.75cm
//		sheet.setColumnWidth(6, (int) (3.75 * 37)); // 3.75cm
//		sheet.setColumnWidth(7, (int) (3.75 * 37)); // 3.75cm
//
//		// create a new font
//		Font font = workbook.createFont();
//		font.setFontHeightInPoints((short) fontSize);
//		font.setFontName(fontName);
//
//		try {
//			FileOutputStream outputStream = new FileOutputStream(outputFilePath);
//			workbook.write(outputStream);
//			workbook.close();
//		} catch (Exception e) {
//			e.printStackTrace();
//		}
//		return workbook;
//	}
//
//	public static void writeLinkedListToExcel(LinkedList<String[]> data, String fileName, String sheetName, XSSFWorkbook workbook) {
//		Sheet sheet = workbook.createSheet(sheetName);
//		int rowNum = 0;
//		for (String[] rowData : data) {
//			Row row = sheet.createRow(rowNum++);
//			int colNum = 0;
//			for (String cellValue : rowData) {
//				Cell cell = row.createCell(colNum++);
//				cell.setCellValue(cellValue);
//			}
//		}
//
//		try {
//			FileOutputStream outputStream = new FileOutputStream(fileName);
//			workbook.write(outputStream);
//			workbook.close();
//		} catch (Exception e) {
//			e.printStackTrace();
//		}
//	}

	public static XSSFWorkbook createExcelSheetWithCellWidth(XSSFWorkbook workbook, boolean closeWorkbookAfterWriting, LinkedList<String[]> data, String outputFilePath,
			String sheetName, String fontName, short fontSize) {
		if (workbook == null) {
			workbook = new XSSFWorkbook();
		}

		Sheet sheet = workbook.createSheet(sheetName);

		// Set the width of each cell in centimeters
//		sheet.setColumnWidth(0, (int) (3.75 * 37)); // 3.75cm
//		sheet.setColumnWidth(1, (int) (2.00 * 37)); // 2.00cm
//		sheet.setColumnWidth(2, (int) (0.75 * 37)); // 0.75cm
//		sheet.setColumnWidth(3, (int) (2.00 * 37)); // 2.00cm
//		sheet.setColumnWidth(4, (int) (0.75 * 37)); // 0.75cm
//		sheet.setColumnWidth(5, (int) (3.75 * 37)); // 3.75cm
//		sheet.setColumnWidth(6, (int) (3.75 * 37)); // 3.75cm
//		sheet.setColumnWidth(7, (int) (3.75 * 37)); // 3.75cm

		// create a new font
		Font font = workbook.createFont();
		font.setFontHeightInPoints((short) fontSize);
		font.setFontName(fontName);
		int rowNum = 0;
		for (String[] rowData : data) {
			Row row = sheet.createRow(rowNum++);
			int colNum = 0;
			for (String cellValue : rowData) {
				Cell cell = row.createCell(colNum++);
				cell.setCellValue(cellValue);
			}
		}
		try {
			FileOutputStream outputStream = new FileOutputStream(outputFilePath);
			workbook.write(outputStream);
			if (closeWorkbookAfterWriting) {
				workbook.close();
			}

		} catch (Exception e) {
			e.printStackTrace();
		}
		return workbook;
	}

//	public static void writeLinkedListToExcel(LinkedList<String[]> data, String filePath, XSSFWorkbook workbook) {
////		Sheet sheet = workbook.createSheet(sheetName);
////		int rowNum = 0;
////		for (String[] rowData : data) {
////			Row row = sheet.createRow(rowNum++);
////			int colNum = 0;
////			for (String cellValue : rowData) {
////				Cell cell = row.createCell(colNum++);
////				cell.setCellValue(cellValue);
////			}
//
////		}
//
//		try {
//			FileOutputStream outputStream = new FileOutputStream(filePath);
//			workbook.write(outputStream);
//			workbook.close();
//		} catch (Exception e) {
//			e.printStackTrace();
//		}
//	}

}
