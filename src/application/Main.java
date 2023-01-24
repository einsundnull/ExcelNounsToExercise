package application;

import java.util.LinkedList;

import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import javafx.application.Application;
import javafx.fxml.FXMLLoader;
import javafx.scene.Scene;
import javafx.scene.layout.AnchorPane;
import javafx.stage.Stage;

public class Main extends Application {

	@Override
	public void start(Stage primaryStage) {
		try {
			setStartValues();
			AnchorPane root = (AnchorPane) FXMLLoader.load(getClass().getResource("Sample.fxml"));
			Scene scene = new Scene(root, 400, 400);
			scene.getStylesheets().add(getClass().getResource("application.css").toExternalForm());
			primaryStage.setScene(scene);
			primaryStage.show();
//			File file = GetFileFromDesktop.getFileFromDesktop();
//			System.out.print(file);

			// Reads the Excel Sheet
			LinkedList<String[]> dataUnsplitted = ExcelReader.readExcelSingleSheet(null);
			// Splits all the Strings into their own column.
			LinkedList<String[]> dataSplitted = ExcelReader.processDataAsLinkedListSplittWhitespace(dataUnsplitted);
			// Creates the first Exercise
			LinkedList<String[]> dataExerciseOne = ExcelReader.processDataToExerciseOne(dataSplitted);
			// Creates the second Exercise
			LinkedList<String[]> dataExerciseTwo = ExcelReader.processDataToExerciseTwo(dataUnsplitted);
			// Here I can make a single String out of all the data.
//			String dataIII = ExcelReader.processDataToSingleString(dataUnsplitted);
//			// Here I copy it to the clipboard
//			ExcelReader.wordsToClipboard(dataIII);
//			for (int i = 0; i < 5; i++) {
//				dataExerciseOne.add(new String[4]);
//			}
//			for (int i = 0; i < dataExerciseTow.size(); i++) {
//				dataExerciseOne.add(dataExerciseTow.get(i));
//			}
			XSSFWorkbook workbook = ExcelReader.createExcelSheetWithCellWidth(null, false, dataUnsplitted, ExcelParameter.excelFilePathOut, ExcelParameter.excelSheetNameOne,
					ExcelParameter.fontNameExcel, (short) 0);
			ExcelReader.createExcelSheetWithCellWidth(workbook, false, dataExerciseOne, ExcelParameter.excelFilePathOut, ExcelParameter.excelSheetNameTwo,
					ExcelParameter.fontNameExcel, (short) 0);
			ExcelReader.createExcelSheetWithCellWidth(workbook, true, dataExerciseTwo, ExcelParameter.excelFilePathOut, ExcelParameter.excelSheetNameThree,
					ExcelParameter.fontNameExcel, (short) 0);

//			sysoList(dataUnsplitted);
//			sysoList(dataExerciseOne);
//			sysoList(dataExerciseTwo);

		} catch (Exception e) {
			e.printStackTrace();
		}
	}

	private void setStartValues() {
		ExcelParameter.sheet = 4;

	}

	public static void main(String[] args) {
		launch(args);
	}

	private void sysoList(LinkedList<String[]> data) {
		for (int i = 0; i < data.size(); i++) {
			for (int n = 0; n < data.get(i).length; n++) {
				System.out.print(data.get(i)[n] + " ");
			}
			System.out.println();
		}
	}
}
