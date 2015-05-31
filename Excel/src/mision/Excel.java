package mision;

import org.apache.poi.hssf.usermodel.HSSFRichTextString;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;

import java.io.FileOutputStream;

/**
 *
 * Created by Intel on 30/05/2015.
 */
public class Excel {
    public static void main() {
        Workbook wb = new HSSFWorkbook();
        Sheet hoja1 = wb.createSheet("nueva hoja");
        Row fila = hoja1.createRow(1);
        Row fila1 = hoja1.createRow(2);
        Row fila2 = hoja1.createRow(3);
        Cell celda1 = fila.createCell(5);
        Cell celda2 = fila.createCell(6);
        Cell celda3 = fila1.createCell(5);
        Cell celda4 = fila1.createCell(6);
        Cell celda5 = fila2.createCell(5);
        Cell celda6 = fila2.createCell(6);
        HSSFRichTextString texto1 = new HSSFRichTextString("Entradas");
        celda1.setCellValue(texto1);
        HSSFRichTextString texto2 = new HSSFRichTextString("Salidas");
        celda2.setCellValue(texto2);
        HSSFRichTextString texto3 = new HSSFRichTextString("14:25");
        celda3.setCellValue(texto3);
        HSSFRichTextString texto4 = new HSSFRichTextString("16:50");
        celda4.setCellValue(texto4);
        HSSFRichTextString texto5 = new HSSFRichTextString("15:10");
        celda5.setCellValue(texto5);
        HSSFRichTextString texto6 = new HSSFRichTextString("17:15");
        celda6.setCellValue(texto6);

        try{
            FileOutputStream Archivo = new FileOutputStream("Horarios.xls");
            wb.write(Archivo);
            Archivo.close();
            System.out.println("Horarios.xls ha sido guardado en el disco.");
        } catch (Exception e) {
            e.printStackTrace();
        }
    }
}
