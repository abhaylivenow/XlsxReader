import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Iterator;
import java.util.List;

public class Main {
    public static void main(String[] args) throws IOException {
        File file = new File("/Users/abhaymaurya/Downloads/hindiString.xlsx");
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
                    default: outputList.add("N/A");
                }
            }
        }
        for(int i = 0 ; i < outputList.size()-1; i+=2) {
            if(!outputList.get(i+1).equals("N/A")) {
                System.out.println("<string name=\"" + outputList.get(i) + "\">" + outputList.get(i + 1) + "</string>");
            }
        }
    }
}