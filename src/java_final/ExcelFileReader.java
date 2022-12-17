package java_final;


import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileInputStream;
import java.io.IOException;

public class Reader {
    public static void main(String[] args) throws IOException {


       // public String[][] readXl ( String) throws {

            String path = "C:\\Users\\gutii\\IdeaProjects\\sami_java_class\\src\\excel_file\\ExcelRead.xlsx";


            FileInputStream fis = new FileInputStream(path);
            XSSFWorkbook wb = new XSSFWorkbook(fis);
            XSSFSheet ws = wb.getSheet("sheet1");

            int rows = ws.getLastRowNum() + 1;
            int cols = ws.getRow(0).getLastCellNum();

            String[][] data = new String[rows][cols];

        for (int r = 0; r < rows; r++){

                XSSFRow myRow = ws.getRow(r);

                for (int c = 0; c < cols; c++) {

                    XSSFCell cell = myRow.getCell(c);

                    String value =myRow.getCell(c).toString();

                    data[r][c] = value;

                    System.out.print(data[r][c] + "\t\t");

                }
                System.out.println();

            }
            fis.close();
            wb.close();


            //return data;
        }

        public static String getCellValue ( Cell cellValue){
            Object value = null;

            switch (cellValue.getCellType()) {
                case STRING:
                    value = cellValue.getStringCellValue();
                    break;
                case NUMERIC:
                    value = cellValue.getNumericCellValue();
                    break;
                case BOOLEAN:
                    value = cellValue.getBooleanCellValue();
                    break;
            }
            return value.toString();

        }

    }
