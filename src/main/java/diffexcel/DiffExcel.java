package diffexcel;

import java.io.File;
import java.io.FileInputStream;
import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class DiffExcel	{

	static String sheetName = new String();
	static int sheetNum = 0;

	public static void main(String[] args)	{
		try	{
			FileInputStream db1 = new FileInputStream(new File(args[0]));
			FileInputStream db2 = new FileInputStream(new File(args[1]));

			XSSFWorkbook wb1 = new XSSFWorkbook(db1);
			XSSFWorkbook wb2 = new XSSFWorkbook(db2);

			for (int i = 0; i < wb1.getNumberOfSheets(); i++)	{

				sheetName = wb1.getSheetName(i);
				sheetNum = i;

				if (compareSheets(wb1.getSheetAt(i), wb2.getSheetAt(i)))	{
					System.out.println("\nThe two sheets are equal.");
				}

				else	{
					System.out.println("\nThe two sheets are different.");
				}

				System.out.println("============================================================");
			}

			db1.close();
			db2.close();
		}
		catch(Exception e)	{
			e.printStackTrace();
		}
	}
	static boolean compareSheets(XSSFSheet s1, XSSFSheet s2)	{

		boolean equalSheets = true;
		
		XSSFRow r0 = s1.getRow(0);

		for (int i = s1.getFirstRowNum(); i <= s1.getLastRowNum(); i++)	{

			XSSFRow r1 = s1.getRow(i);
			XSSFRow r2 = s2.getRow(i);

			if (!(compareRows(r0, r1, r2)))	{
				equalSheets = false;
				System.out.println("Row number: "+(i+1));
				System.out.println("Sheet name: "+sheetName);
				System.out.println("---------------------------");
			}
		}
		return equalSheets;
	}
	static boolean compareRows(XSSFRow r0, XSSFRow r1, XSSFRow r2)	{
		if ((r1 == null) && (r2 == null))	{
			return true;
		}
		else if ((r1 == null) || (r2 == null))	{
			return false;
		}

		boolean equalRows = true;

		for (int i = 0; i <= 31; i++)	{

			XSSFCell c1 = r1.getCell(i);
			XSSFCell c2 = r2.getCell(i);
			
			if (!(compareCells(c1, c2)))	{
				equalRows = false;
				XSSFCell cellName = r0.getCell(i);
				XSSFCell food = r1.getCell(2);

				if (sheetNum == 1)
					food = r1.getCell(0);

				cellName.setCellType(CellType.STRING);
				food.setCellType(CellType.STRING);

				System.out.println("Food: "+food.getStringCellValue());
				System.out.println("Cell name: "+cellName.getStringCellValue());
				System.out.println("Cell number: "+(i+1));
				//break;
			}
		}
		return equalRows;
	}
	static boolean compareCells(XSSFCell c1, XSSFCell c2)	{

		if ((c1 == null) && (c2 == null))	{
			return true;
		}

		else if ((c1 == null) && (c2 != null))	{
			c2.setCellType(CellType.STRING);
			System.out.println("C1: \nC2: "+c2.getStringCellValue());
			return false;
		}

		else if ((c1 != null) && (c2 == null))	{
			c1.setCellType(CellType.STRING);

			if (c1.getStringCellValue().equals(""))
				return true;

			System.out.println("C1: "+c1.getStringCellValue()+"\nC2: ");
			return false;
		}

		else if ((c1 != null) && (c2 != null))	{
			c1.setCellType(CellType.STRING);
			c2.setCellType(CellType.STRING);
			if (c1.getStringCellValue().equals(c2.getStringCellValue()))
				return true;
			else	{
				System.out.println("C1: "+c1.getStringCellValue()+"\nC2: "+c2.getStringCellValue());
				return false;
			}
		}
		return false;
	}
}
