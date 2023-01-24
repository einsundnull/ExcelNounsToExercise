package application;

import java.io.File;
import java.util.Scanner;

import javax.swing.JFileChooser;

public class GetFileFromDesktop {

    public static File getFileFromDesktop() {
        // Create a file chooser
        JFileChooser chooser = new JFileChooser();
        chooser.setCurrentDirectory(new File(System.getProperty("user.home") + "/Desktop/Nouns.xlsx"));
        // Set the file selection to only show files, not directories
        chooser.setFileSelectionMode(JFileChooser.FILES_ONLY);
        // Show the file chooser and wait for the user to select a file
        int result = chooser.showOpenDialog(null);
        if (result == JFileChooser.APPROVE_OPTION) {
            // Return the selected file
            return chooser.getSelectedFile();
        } else {
            // If the user did not select a file, return null
            return null;
        }
    }
    

}
