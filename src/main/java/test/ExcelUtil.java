package test;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;

public class ExcelUtil {
    public static void main(String[] args) throws IOException {
        ClassLoader classLoader = ExcelUtil.class.getClassLoader();

        File file = new File(classLoader.getResource("bn-arv-dung-tld-template.xlsx").getFile());

        FileInputStream inputStream =  new FileInputStream(file);
        Workbook workbook = new XSSFWorkbook(inputStream);
        Sheet sheet = workbook.getSheetAt(0);
        Row row = sheet.createRow(1);
        Cell cell = row.createCell(0);
        cell.setCellValue("oke ment");
        FileOutputStream outputStream = new FileOutputStream("bn-arv-dung-tld-template.xlsx");
        workbook.write(outputStream);
        System.out.println("oke");
    }
}
