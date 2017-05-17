package com.fjhw.test;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.io.InputStream;
import java.util.Map;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.junit.Test;

/** ����xls��ʽ��excel */
public class TestReadXLS {
	@Test
	public void test() {
		try {
			InputStream is = new FileInputStream("(������)����������Ϣ����վ����ƽ̨���ݵ����b.xls");
			HSSFWorkbook hssfWorkbook = new HSSFWorkbook(is);
			// ��ȡÿһ��������
			for (int numSheet = 0; numSheet < hssfWorkbook.getNumberOfSheets(); numSheet++) {
				HSSFSheet hssfSheet = hssfWorkbook.getSheetAt(numSheet);
				if (hssfSheet == null) {
					continue;
				}
				// ��ȡ������������
				String sheetName = hssfSheet.getSheetName();
				System.out.println("���빤����������:" + sheetName);
				// ͨ�����к��еķ�ʽ���ʹ������ĵ�Ԫ��
				for (int rowNum = 0; rowNum <= hssfSheet.getLastRowNum(); rowNum++) {
					HSSFRow hssfRow = hssfSheet.getRow(rowNum);
					if (hssfRow == null)
						continue;
					// ��ȡָ�������еı��
					if (hssfRow != null) {
						HSSFCell one = hssfRow.getCell(0);
						// System.out.print(getValue(one));
						// ��ȡ��һ������
						HSSFCell two = hssfRow.getCell(1);
						// System.out.println(getValue(two));
						// ��ȡ�ڶ�������
						HSSFCell three = hssfRow.getCell(2);
						// System.out.println(getValue(three));
						// ��ȡ����������
					}
					// ������ȡ���м��еı��
					int lastCellNum = hssfRow.getLastCellNum();
					for (int i = 0; i < lastCellNum; i++) {
						HSSFCell cell = hssfRow.getCell(i);
						if (cell == null)
							continue;
						System.out.print(getValue(cell) + "  ");
					}
					System.out.println();
				}
			}
		} catch (FileNotFoundException e) {
			e.printStackTrace();
		} catch (IOException e) {
			e.printStackTrace();
		}
	}

	private static String getValue(HSSFCell hssfCell) {
		if (hssfCell.getCellType() == HSSFCell.CELL_TYPE_BOOLEAN) {
			return String.valueOf(hssfCell.getBooleanCellValue());
		} else if (hssfCell.getCellType() == HSSFCell.CELL_TYPE_NUMERIC) {
			return String.valueOf(hssfCell.getNumericCellValue());
		} else if (hssfCell.getCellType() == HSSFCell.CELL_TYPE_STRING) {
			return String.valueOf(hssfCell.getStringCellValue());
		} else {
			return null;
		}
	}
}
