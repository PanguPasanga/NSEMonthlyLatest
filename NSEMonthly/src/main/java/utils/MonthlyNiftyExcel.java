package utils;

import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.List;

public class MonthlyNiftyExcel {
    public static int k;

    public static void writeToExcel(List<String> optionValues, String strikeValues) throws IOException {
        String file = ".\\DataSheet\\Nifty_Monthly_Nov.xlsx";
        FileInputStream fis = new FileInputStream(file);
        XSSFWorkbook wb = new XSSFWorkbook(fis);
        XSSFSheet sheet;
        String strikeData = strikeValues;
        switch (strikeData) {
            case "10700.00":
                sheet = wb.getSheet("10700");
                break;
            case "10800.00":
                sheet = wb.getSheet("10800");
                break;
            case "10900.00":
                sheet = wb.getSheet("10900");
                break;
            case "11000.00":
                sheet = wb.getSheet("11000");
                break;
            case "11100.00":
                sheet = wb.getSheet("11100");
                break;
            case "11200.00":
                sheet = wb.getSheet("11200");
                break;
            case "11300.00":
                sheet = wb.getSheet("11300");
                break;
            case "11400.00":
                sheet = wb.getSheet("11400");
                break;
            case "11500.00":
                sheet = wb.getSheet("11500");
                break;
            case "11600.00":
                sheet = wb.getSheet("11600");
                break;
            case "11700.00":
                sheet = wb.getSheet("11700");
                break;
            default:
                throw new IllegalStateException("Unexpected value: " + strikeData);
        }
        //int rowCount=sheet.getLastRowNum();
        int colCount = sheet.getRow(0).getLastCellNum();

        int index = 0;
        for (int i = k; i <= 30; i++) {
            Row rw = sheet.getRow(i);
            if (rw == null) {
                    rw= sheet.createRow(i);
                for (int j = 0; j < colCount; j++) {
                    //XSSFCell cw = rw.getCell(j);
                    rw.createCell(j).setCellValue(optionValues.get(index++));
                   /* try {
                        if (cw == null) {
                            rw.createCell(j).setCellValue(optionValues.get(index++));

                        }

                    } catch (Exception e) {
                        System.out.println(e);
                    }*/
                }
            }
            FileOutputStream fos = new FileOutputStream(file);
            wb.write(fos);
            fos.close();
            wb.close();
            k++;
            System.out.println("the value of K: " + k);

        }


    }
}

