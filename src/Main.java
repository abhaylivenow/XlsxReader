import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import javax.swing.*;
import java.awt.event.ActionEvent;
import java.awt.event.ActionListener;
import java.io.*;
import java.util.ArrayList;
import java.util.Iterator;
import java.util.List;

public class Main {
    public static void main(String[] args) {

        JFrame frame = new JFrame();
        JLabel labelSource = new JLabel("Source Location");
        JTextField edtSourceLocation = new JTextField(20);
        JLabel labelDestination = new JLabel("Destination Location");
        JTextField edtDestinationLocation = new JTextField(20);
        JButton btnExcelToXml = new JButton("Generate xml files");
        JPanel p1 = new JPanel();
        p1.add(labelSource);
        p1.add(edtSourceLocation);
        p1.add(labelDestination);
        p1.add(edtDestinationLocation);
        p1.add(btnExcelToXml);

        JPanel p2 = new JPanel();

        JTabbedPane tabs = new JTabbedPane();
        tabs.setBounds(40,20,300,300);
        tabs.add("Excel to xml", p1);
        tabs.add("Xml to excel", p2);

        //Add tabs to the frame
        frame.add(tabs);

        frame.setSize(400,400);
        frame.setLayout(null);
        frame.setVisible(true);

        btnExcelToXml.addActionListener(e -> {
            System.out.println(edtSourceLocation.getText());
            System.out.println(edtDestinationLocation.getText());
            try {
                init(edtSourceLocation.getText(),edtDestinationLocation.getText());
            } catch (IOException ex) {
                throw new RuntimeException(ex);
            }
        });
    }

    private static void init(String sourceLocation, String destinationLocation) throws IOException {
        File file = new File(sourceLocation);
        FileInputStream fis = new FileInputStream(file);
        XSSFWorkbook wb = new XSSFWorkbook(fis);
        XSSFSheet sheet = wb.getSheetAt(0);
        Iterator<Row> itr = sheet.iterator();
        List<String> outputList = new ArrayList<>();
        while (itr.hasNext()) {
            Row row = itr.next();
            Iterator<Cell> cellIterator = row.cellIterator();
            while (cellIterator.hasNext()) {
                Cell cell = cellIterator.next();
                switch (cell.getCellType()) {
                    case Cell.CELL_TYPE_STRING:
                        outputList.add(cell.getStringCellValue());
                        break;
                    case Cell.CELL_TYPE_NUMERIC:
                        outputList.add(String.valueOf(cell.getNumericCellValue()));
                        break;
                    default:
                        outputList.add("N/A");
                }
            }
        }

        generateFilesAndFolders(outputList, destinationLocation);
    }

    private static void generateFilesAndFolders(List<String> outputList, String destinationLocation) throws IOException {
        File f1 = new File(destinationLocation+"GeneratedFolders");
        f1.mkdir();
        createEnglishFile(outputList, destinationLocation);
        createHindiFile(outputList, destinationLocation);
        createHinglishFile(outputList, destinationLocation);
    }

    private static void createEnglishFile(List<String> outputList, String destinationLocation) throws IOException {
        File file = new File(destinationLocation+"GeneratedFolders/EnglishFile.txt");
        boolean result;
        try {
            result = file.createNewFile();
            if (result) {
                System.out.println("file created " + file.getCanonicalPath());
            } else {
                System.out.println("File already exist at location: " + file.getCanonicalPath());
            }
        } catch (IOException e) {
            e.printStackTrace();
        }

        FileWriter writer = new FileWriter(destinationLocation+"GeneratedFolders/EnglishFile.txt");

        try {
            for (int i = 0; i < outputList.size() - 1; i += 4) {
                if (!outputList.get(i + 1).equals("N/A")) {
                    writer.write("<string name=\"" + outputList.get(i) + "\">" + outputList.get(i + 1) + "</string> \n");
                }
            }
        } catch (Exception e) {
            e.printStackTrace();
        }
    }

    private static void createHinglishFile(List<String> outputList, String destinationFolder) throws IOException {
        File file = new File(destinationFolder+"GeneratedFolders/HinglishFile.txt");
        boolean result;
        try {
            result = file.createNewFile();
            if (result) {
                System.out.println("file created " + file.getCanonicalPath());
            } else {
                System.out.println("File already exist at location: " + file.getCanonicalPath());
            }
        } catch (IOException e) {
            e.printStackTrace();
        }

        FileWriter writer = new FileWriter(destinationFolder+"GeneratedFolders/HinglishFile.txt");

        try {
            for (int i = 0; i < outputList.size() - 1; i += 4) {
                if (!outputList.get(i + 1).equals("N/A")) {
                    writer.write("<string name=\"" + outputList.get(i) + "\">" + outputList.get(i + 2) + "</string> \n");
                }
            }
        } catch (Exception e) {
            e.printStackTrace();
        }
    }

    private static void createHindiFile(List<String> outputList, String destinationFolder) throws IOException {
        File file = new File(destinationFolder+"GeneratedFolders/HindiFile.txt");
        boolean result;
        try {
            result = file.createNewFile();
            if (result) {
                System.out.println("file created " + file.getCanonicalPath());
            } else {
                System.out.println("File already exist at location: " + file.getCanonicalPath());
            }
        } catch (IOException e) {
            e.printStackTrace();
        }

        FileWriter writer = new FileWriter(destinationFolder+"GeneratedFolders/HindiFile.txt");

        try {
            for (int i = 0; i < outputList.size() - 1; i += 4) {
                if (!outputList.get(i + 1).equals("N/A")) {
                    writer.write("<string name=\"" + outputList.get(i) + "\">" + outputList.get(i + 3) + "</string> \n");
                }
            }
        } catch (Exception e) {
            e.printStackTrace();
        }
    }
}