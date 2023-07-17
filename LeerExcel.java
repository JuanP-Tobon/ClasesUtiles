package co.com.devco.models;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.*;



public class LeerExcel {

    private LeerExcel(){}

    public static String retornarValorExcel(int numeroFila, int numeroColumna, String rutaExcel) {

        File f = new File(rutaExcel);

        try {
            InputStream inp = new FileInputStream(f);
            Workbook wb = WorkbookFactory.create(inp);
            Sheet sheet = wb.getSheetAt(0);

            Row row = sheet.getRow(numeroFila);
            Cell cell = row.getCell(numeroColumna);
            String value = cell.getStringCellValue();
            return value;
        } catch (FileNotFoundException e) {
            throw new RuntimeException(e);
        } catch (IOException e) {
            throw new RuntimeException(e);
        }
    }

    public static void cambiarValor(int numeroFila, int numeroColumna, String valor, String rutaExcel) throws FileNotFoundException {



        try {
            FileInputStream file = new FileInputStream(rutaExcel);
            XSSFWorkbook wb = new XSSFWorkbook(file);
            XSSFSheet sheet = wb.getSheetAt(0);

            XSSFRow fila = sheet.getRow(numeroFila);
            XSSFCell cell = fila.getCell(numeroColumna);

            cell.setCellValue(valor);
            file.close();

            FileOutputStream output = new FileOutputStream(rutaExcel);
            wb.write(output);
            output.close();

        } catch (IOException e) {
            throw new RuntimeException(e);
        }

    }
}
