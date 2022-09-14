import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileWriter;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Iterator;
import java.util.List;

public class Main {
    public static void main(String[] args) throws IOException {
        File file = new File("/Users/abhaymaurya/Downloads/string.xlsx");
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

        generateFiles(outputList);
    }

    private static void generateFiles(List<String> outputList) throws IOException {
        createEnglishFile(outputList);
        createHindiFile(outputList);
        createHinglishFile(outputList);
    }

    private static void createEnglishFile(List<String> outputList) throws IOException {
        File file = new File("/Users/abhaymaurya/Downloads/EnglishFile.txt");
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

        FileWriter writer = new FileWriter("/Users/abhaymaurya/Downloads/EnglishFile.txt");

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

    private static void createHinglishFile(List<String> outputList) throws IOException {
        File file = new File("/Users/abhaymaurya/Downloads/HinglishFile.txt");
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

        FileWriter writer = new FileWriter("/Users/abhaymaurya/Downloads/HinglishFile.txt");

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

    private static void createHindiFile(List<String> outputList) throws IOException {
        File file = new File("/Users/abhaymaurya/Downloads/HindiFile.txt");
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

        FileWriter writer = new FileWriter("/Users/abhaymaurya/Downloads/HindiFile.txt");

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