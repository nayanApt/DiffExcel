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
	public static void main(String[] args)	{
		try	{
			FileInputStream db1 = new FileInputStream(new File(args[0]));
			FileInputStream db2 = new FileInputStream(new File(args[1]));

			XSSFWorkbook wb1 = new XSSFWorkbook(db1);
			XSSFWorkbook wb2 = new XSSFWorkbook(db2);

			for (int i = 0; i < wb1.getNumberOfSheets(); i++)	{

				System.out.println("Comparing sheet number: "+(i+1));

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

		for (int i = s1.getFirstRowNum(); i <= s1.getLastRowNum(); i++)	{

			XSSFRow r1 = s1.getRow(i);
			XSSFRow r2 = s2.getRow(i);

			if (!(compareRows(r1, r2)))	{
				equalSheets = false;
				System.out.println("Rows "+(i+1)+" NOT equal.");
				System.out.println("---------------------------");
			}
		}
		return equalSheets;
	}
	static boolean compareRows(XSSFRow r1, XSSFRow r2)	{
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
				System.out.println("\tCells "+(i+1)+" NOT equal.");
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
			System.out.println("\tC1: \n\tC2: "+c2.getStringCellValue());
			return false;
		}

		else if ((c1 != null) && (c2 == null))	{
			c1.setCellType(CellType.STRING);

			if (c1.getStringCellValue().equals(""))
				return true;

			System.out.println("\tC1: "+c1.getStringCellValue()+"\n\tC2: ");
			return false;
		}

		else if ((c1 != null) && (c2 != null))	{
			c1.setCellType(CellType.STRING);
			c2.setCellType(CellType.STRING);
			if (c1.getStringCellValue().equals(c2.getStringCellValue()))
				return true;
			else	{
				System.out.println("\tC1: "+c1.getStringCellValue()+"\n\tC2: "+c2.getStringCellValue());
				return false;
			}
		}
		return false;
	}
}
