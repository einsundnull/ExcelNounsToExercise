package application;



	import java.util.LinkedList;

import javafx.application.Application;
import javafx.collections.FXCollections;
import javafx.collections.ObservableList;
import javafx.geometry.Insets;
import javafx.scene.Scene;
import javafx.scene.control.Label;
import javafx.scene.control.ListView;
import javafx.scene.control.TextField;
import javafx.scene.layout.HBox;
import javafx.scene.layout.VBox;
import javafx.stage.Stage;

	public class ListViewCards extends Application {
	  
		private LinkedList<String[]> data;

	    public ListViewCards(LinkedList<String[]> data) {
	        this.data = data;
	    }

	    @Override
	    public void start(Stage primaryStage) {
	        // Create the ListView
	        ListView<HBox> listView = new ListView<>();
	        listView.setPrefWidth(400);
	        listView.setPrefHeight(200);
//	        data = ExcelReader.readExcel(null);
	        // Create an observable list to hold the items in the ListView
	        ObservableList<HBox> items = FXCollections.observableArrayList();

	        // Add the items to the observable list
	        for (String[] cells : data) {
	            HBox hbox = new HBox();
	            hbox.setSpacing(10);
	            hbox.setPadding(new Insets(10));
	            for (String cell : cells) {
	                TextField textField = new TextField(cell);
	                textField.setEditable(false);
	                hbox.getChildren().add(textField);
	            }
	            items.add(hbox);
	        }
	        listView.setItems(items);

	        // Create a root container to hold the ListView
	        VBox root = new VBox();
	        root.setPadding(new Insets(10));
	        root.getChildren().addAll(new Label("Data"), listView);

	        // Create a scene and set it to the stage
	        Scene scene = new Scene(root);
	        primaryStage.setScene(scene);
	        primaryStage.show();
	    }
	    
	    public static void main(String[] args) {
	        LinkedList<String[]> data = ExcelReader.readExcelAllSheets(null);
	        launch(args);
	    }
	}
