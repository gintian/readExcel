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

/** 解析xls格式的excel */
public class TestReadXLS {
	@Test
	public void test() {
		try {
			InputStream is = new FileInputStream("(崇仁县)乡镇气象信息服务站工作平台数据调查表b.xls");
			HSSFWorkbook hssfWorkbook = new HSSFWorkbook(is);
			// 获取每一个工作薄
			for (int numSheet = 0; numSheet < hssfWorkbook.getNumberOfSheets(); numSheet++) {
				HSSFSheet hssfSheet = hssfWorkbook.getSheetAt(numSheet);
				if (hssfSheet == null) {
					continue;
				}
				// 获取工作簿的名字
				String sheetName = hssfSheet.getSheetName();
				System.out.println("输入工作簿的名字:" + sheetName);
				// 通过先行后列的方式访问工作簿的单元格
				for (int rowNum = 0; rowNum <= hssfSheet.getLastRowNum(); rowNum++) {
					HSSFRow hssfRow = hssfSheet.getRow(rowNum);
					if (hssfRow == null)
						continue;
					// 读取指定的行列的表格
					if (hssfRow != null) {
						HSSFCell one = hssfRow.getCell(0);
						// System.out.print(getValue(one));
						// 读取第一列数据
						HSSFCell two = hssfRow.getCell(1);
						// System.out.println(getValue(two));
						// 读取第二列数据
						HSSFCell three = hssfRow.getCell(2);
						// System.out.println(getValue(three));
						// 读取第三列数据
					}
					// 遍历读取几行几列的表格
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
