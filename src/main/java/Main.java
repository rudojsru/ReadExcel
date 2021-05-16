
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.FormulaEvaluator;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.util.ArrayList;
import java.util.List;

public class Main {
    static String FILE = "C:/Users/Olexandr.Rudyy/Desktop/FN.xlsx";
    static FormulaEvaluator formulaEvaluator;

    public static void main(String[] args) {
        ReadWreitExcel readWreitExcel = new ReadWreitExcel();
        XSSFWorkbook wb = readWreitExcel.openBook(FILE);
        XSSFSheet sheet = wb.getSheet("Februar");
        XSSFSheet sheet1 = wb.getSheet("Marz");
        XSSFSheet sheet2 = wb.getSheet("April");
        List list = new ArrayList();
        formulaEvaluator = wb.getCreationHelper().createFormulaEvaluator();
//1
        for (Row row : sheet) {
            String temp = addValue(row);
            list.add(temp);
        }
//2
        int counter = 0;
        for (int i = 1; i <= sheet1.getLastRowNum(); i++) {
            String temp = addValue(sheet1.getRow(i));
            if (list.contains(temp)) {
                sheet1.getRow(i).createCell(14).setCellValue("0");
                //  sheet1.getRow(i).createCell(15).setCellValue(i);
            } else {
                sheet1.getRow(i).createCell(14).setCellValue("1");
                counter++;
                sheet1.getRow(i).createCell(15).setCellValue(counter);
                list.add(temp);
            }
        }
//3
        counter = 0;
        for (int i = 1; i <= sheet2.getLastRowNum(); i++) {
            String temp = addValue(sheet2.getRow(i));
            if (list.contains(temp)) {
                sheet2.getRow(i).createCell(14).setCellValue("0");
            } else {
                sheet2.getRow(i).createCell(14).setCellValue("1");
                counter++;
                sheet2.getRow(i).createCell(15).setCellValue(counter);
                list.add(temp);
            }
        }

        readWreitExcel.writeWorkbook(wb, FILE);
        System.out.println(list.size());
    }

    private static String addValue(Row row) {
        String temp = "";
        switch (formulaEvaluator.evaluateInCell(row.getCell(0)).getCellType()) {
            case Cell.CELL_TYPE_NUMERIC:   //field that represents numeric cell type
                temp = (int) row.getCell(0).getNumericCellValue() + "\t\t";
                break;
            case Cell.CELL_TYPE_STRING:    //field that represents string cell type
                temp = row.getCell(0).getStringCellValue() + "\t\t";
                break;
        }
        return temp;
    }

}