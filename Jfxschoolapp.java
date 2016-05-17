package schoolapp;

import java.awt.print.PrinterException;
import java.io.FileInputStream;
import java.io.IOException;
import java.text.MessageFormat;
import java.util.Iterator;

import javax.swing.JEditorPane;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFFormulaEvaluator;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.FormulaEvaluator;
import org.apache.poi.ss.usermodel.Row;

import javafx.application.Application;
import javafx.event.ActionEvent;
import javafx.scene.Scene;
import javafx.scene.control.Button;
import javafx.scene.control.Label;
import javafx.scene.control.Menu;
import javafx.scene.control.MenuBar;
import javafx.scene.control.MenuItem;
import javafx.scene.control.TextField;
import javafx.scene.effect.InnerShadow;
import javafx.scene.input.KeyCode;
import javafx.scene.input.KeyCodeCombination;
import javafx.scene.input.KeyCombination;
import javafx.scene.layout.AnchorPane;
import javafx.scene.layout.BorderPane;
import javafx.scene.web.WebEngine;
import javafx.scene.web.WebView;
import javafx.stage.Stage;

public class Jfxschoolapp extends Application {

    public static void gettheLesson(String student, int lessonnumber) {
        try {
            FileInputStream inp = new FileInputStream("C:/Spreadsheets/" + student + ".xls");
            HSSFWorkbook workbook = new HSSFWorkbook(inp);
            DataFormatter objDefaultFormat = new DataFormatter();
            FormulaEvaluator objFormulaEvaluator = new HSSFFormulaEvaluator((HSSFWorkbook) workbook);
            HSSFSheet worksheet = workbook.getSheet("Sheet1");
            Iterator<Row> objIterator = worksheet.rowIterator();

            int count = 0;
            String content = "";
            String subject = "";
            // System.out.println(cellValueStr);
            worksheet.getPhysicalNumberOfRows();

            while (objIterator.hasNext()) {

                HSSFRow row = (HSSFRow) objIterator.next();
                HSSFCell cellValue = row.getCell(1);
                // objFormulaEvaluator.evaluate(cellValue); // This will
                // evaluate
                // the cell, And any
                // type of cell will
                // return string
                // value
                String cellValueStr = objDefaultFormat.formatCellValue(cellValue, objFormulaEvaluator);
                // System.out.println(cellValueStr);
                if (cellValueStr.length() > 0) {
                    int cellnumber = Integer.parseInt(cellValueStr);
                    // if the lesson\/==lesson(request) &
                    if (cellnumber == lessonnumber) {
                        HSSFRow r = cellValue.getRow();
                        HSSFCell cell = r.getCell(7);
                        String assignment = cell.getStringCellValue();
                        // System.out.println(assignment);
                        HSSFCell cl = r.getCell(4);
                        String currentsub = cl.getStringCellValue();

                        if (count == 0) {
                            // make the html initially and add the row
                            subject = currentsub;

                            content += "<html><body><hr>\n" + "<u><p  style= \" color: rgb(115,115,115);\""
                                    + ">Subject: " + subject + "</p></u><p  style= \" color: rgb(115,115,115);"
                                    + "\">__________&nbsp;&nbsp;&nbsp; " + assignment + "</p>\n";

                            // System.out.println(subject + currentsub);
                        }
                        if (count > 0) {
                            // IF it is not the first assignment, add the row
                            // (assignment)
                            // subject is the last subject used
                            // currentsub is the subject being written now.
                            if (!subject.equals(currentsub)) {
                                // if the last subject does NOT equal the
                                // current subject
                                content += "<hr>";
                                subject = currentsub;
                                content += "<p  style= \" color: rgb(115,115,115);\"><u>Subject: " + subject
                                        + "</u></p><p  style= \" color: rgb(115,115,115);\""
                                        + ">__________&nbsp;&nbsp;&nbsp; " + assignment + "</p>\n";
                                // reset the last subject to be the current
                                // subject
                                // and write the header and footer
                            } else if (subject.equals(currentsub)) {
                                content += "<p  style= \" color: rgb(115,115,115); \""
                                        + ">__________ &nbsp;&nbsp;&nbsp;" + assignment + "</p>\n";
                            }

                            // System.out.println(count);
                            // System.out.println(subject + currentsub);
                        }
                        count++;// makes the count go up one
                    }
                }
            }
            Stage primaryStage = new Stage();
            primaryStage.setWidth(400);
            primaryStage.setHeight(300);
            primaryStage.setMaximized(true);

            BorderPane root = new BorderPane();
            Scene scene = new Scene(root);
            primaryStage.setScene(scene);

            primaryStage.setTitle(student + "'s Lesson " + lessonnumber);

            MenuBar menuBar = new MenuBar();
            root.setTop(menuBar);

            WebView wv = new WebView();
            WebEngine we = wv.getEngine();
            we.loadContent(content);
            root.setCenter(wv);

            JEditorPane editorPane = new JEditorPane();
            editorPane.setContentType("text/html");
            editorPane.setText(content);
            Menu mnFile = new Menu("File");
            menuBar.getMenus().add(mnFile);
            MenuItem mntmPrint = new MenuItem("Print");
            mntmPrint.setAccelerator(new KeyCodeCombination(KeyCode.P, KeyCombination.CONTROL_DOWN));
            mntmPrint.setOnAction((ActionEvent event) -> {
                try {
                    MessageFormat footer = new MessageFormat("Page - {0}");
                    MessageFormat header = new MessageFormat("Lesson: " + lessonnumber + "     " + student);
                    editorPane.print(header, footer);
                } catch (PrinterException e) {
                }
            });
            mnFile.getItems().add(mntmPrint);
            primaryStage.show();
            //System.gc();
        } catch (IOException | NumberFormatException r) {
        }

    }

    public static void main(String[] args) {
        launch(args);
    }

    @Override
    public void start(Stage primaryStage) throws Exception {
        AnchorPane root = new AnchorPane();
        Scene scene = new Scene(root);
        primaryStage.setScene(scene);
        primaryStage.setHeight(350);
        primaryStage.setWidth(590);

        Label welcomeschapp = new Label("Welcome to the SchoolApp!");
        welcomeschapp.setFont(new javafx.scene.text.Font(25));
        welcomeschapp.setLayoutX(10);
        welcomeschapp.setLayoutY(10);
        root.getChildren().add(welcomeschapp);

        InnerShadow shadow = new InnerShadow();
        shadow.setRadius(20);
        root.setEffect(shadow);

        Label lblLessonnumber = new Label("Finest School in the World!!!");
        lblLessonnumber.setFont(new javafx.scene.text.Font(14));
        lblLessonnumber.setLayoutX(15);
        lblLessonnumber.setLayoutY(45);
        root.getChildren().add(lblLessonnumber);

        TextField lessonbox = new TextField();

        Label lesson = new Label("Lesson:  ");
        lesson.setLayoutX(15);
        lesson.setLayoutY(77);
        lesson.setFont(new javafx.scene.text.Font(14));
        root.getChildren().add(lesson);

        lessonbox.setLayoutY(75);
        lessonbox.setLayoutX(70);
        root.getChildren().add(lessonbox);

        Button btn1 = new Button("Student1");
        root.getChildren().add(btn1);
        btn1.setLayoutX(15);
        btn1.setLayoutY(110);

        Button btn2 = new Button("Student2");
        btn2.setLayoutX(15);
        btn2.setLayoutY(140);
        root.getChildren().add(btn2);

        Button btn3 = new Button("Student3");
        btn3.setLayoutX(15);
        btn3.setLayoutY(170);
        root.getChildren().add(btn3);

        Button btn4 = new Button("Student4");
        btn4.setLayoutX(15);
        btn4.setLayoutY(200);
        root.getChildren().add(btn4);

        Button btn5 = new Button("Student5");
        btn5.setLayoutX(15);
        btn5.setLayoutY(230);
        root.getChildren().add(btn5);
        btn1.setOnAction((ActionEvent event) -> {
            gettheLesson("Student1", Integer.parseInt(lessonbox.getText()));
        });

        btn2.setOnAction((ActionEvent event) -> {
            gettheLesson("Student2", Integer.parseInt(lessonbox.getText()));
        });

        btn3.setOnAction((ActionEvent event) -> {
            gettheLesson("Student3", Integer.parseInt(lessonbox.getText()));
        });

        btn4.setOnAction((ActionEvent event) -> {
            gettheLesson("Student4", Integer.parseInt(lessonbox.getText()));
        });

        btn5.setOnAction((ActionEvent event) -> {
            gettheLesson("Student5", Integer.parseInt(lessonbox.getText()));
        });
        primaryStage.setTitle("School App 3");
        primaryStage.show();

    }
}
