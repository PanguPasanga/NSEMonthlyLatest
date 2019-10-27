package utils;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ReadExcel 
{
	public static void main(String args[]) throws IOException
	{
		File file= new File(".\\DataSheet\\Nifty_Weekly_17Oct.xlsx");
				FileInputStream fis= new FileInputStream(file);
				XSSFWorkbook wb= new XSSFWorkbook(fis);
				XSSFSheet sh= wb.getSheet("Weekly");
				int rowCount =sh.getLastRowNum();
				int cellCount=sh.getRow(0).getLastCellNum();
				System.out.println(rowCount);
				for(int i=0;i<=rowCount;i++)
				{
					XSSFRow rw=sh.getRow(i);
					if(rw==null)
					{
						continue;
					}
					
					for(int j=0;j<cellCount;j++)
					{
						XSSFCell cw=rw.getCell(j);
						if(cw==null)
						{
							if(sh.getRow(i).getCell(0).toString().equalsIgnoreCase("11700.00_SATURDAY"))
							{
								System.out.println(rw.getRowNum()+"---->"+rw.getCell(0).toString());
								rw.createCell(j).setCellValue("Nagarajan");
								
							}
					
						}
					}
				}
				FileOutputStream fos= new FileOutputStream(file);
				wb.write(fos);
				fos.close();
				wb.close();
				
		
	}

}
