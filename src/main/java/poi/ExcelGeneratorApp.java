package poi;

import java.io.BufferedOutputStream;
import java.io.FileOutputStream;
import java.io.OutputStream;
import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.CreationHelper;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

/**
 * This code generates sample excel file with colors and borders for easier style selection with Apache POI lib!
 * @author Valentin Kovrov
 */
public class ExcelGeneratorApp {
    public static void main(String args[]) throws Exception {
        try (
            Workbook workbook = new XSSFWorkbook();
            OutputStream os= new BufferedOutputStream(new FileOutputStream("styles.xlsx"));
                ) {
            CreationHelper createHelper = workbook.getCreationHelper();

            Sheet sheet = workbook.createSheet("Colors");

            int colMax= 5;
            for (int i=0; i< colMax; i++) {
                sheet.setColumnWidth(i, 20*256); // width = chars * 256
            }             
            
            int rowNum = 0, colNum = 0;
            Row row = sheet.createRow(rowNum++);
            for (IndexedColors color : IndexedColors.values()) {
                CellStyle style = workbook.createCellStyle();
                style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
                style.setFillForegroundColor(color.getIndex());                

                if (colNum>=colMax) {
                    row = sheet.createRow(rowNum++);
                    colNum= 0;
                }
                
                Cell cell = row.createCell(colNum++);
                cell.setCellValue(color.name());
                cell.setCellStyle(style);                
            }
            
            rowNum++;
            colNum= 0;
            row = sheet.createRow(rowNum++);
            for (BorderStyle border : BorderStyle.values()) {
                CellStyle style = workbook.createCellStyle();

                style.setBorderBottom(border);
                style.setBorderLeft(border);
                style.setBorderRight(border);
                style.setBorderTop(border);                    
                
                
                if (colNum+1>=colMax) {
                    rowNum++;
                    row = sheet.createRow(rowNum++);
                    colNum= 0;
                }
                
                colNum++;
                Cell cell = row.createCell(colNum++);
                cell.setCellValue(border.name());
                cell.setCellStyle(style);                
                
            }
                        
            workbook.write(os);
        }
    }    
}
