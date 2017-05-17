package com.fjhw.test;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.io.InputStream;

import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.Test;

/**解析xlsx格式的excel*/
public class TestReadXLSX {
	@Test
	public void test1(){
		try {
			InputStream is = new FileInputStream("东乡：乡镇气象信息服务站工作平台数据调查表.xlsx");
			XSSFWorkbook xssfWorkbook = new XSSFWorkbook(is);
			// 获取每一个工作薄
			for (int numSheet = 0; numSheet < xssfWorkbook.getNumberOfSheets(); numSheet++) {
			    XSSFSheet xssfSheet = xssfWorkbook.getSheetAt(numSheet);
			    if (xssfSheet == null) {
			        continue;
			    }
			    // 获取当前工作薄的每一行
			    for (int rowNum = 1; rowNum <= xssfSheet.getLastRowNum(); rowNum++) {
			        XSSFRow xssfRow = xssfSheet.getRow(rowNum);
			        if (xssfRow != null) {
			        	//读取第一列数据
			            XSSFCell one = xssfRow.getCell(0);
			            if(one==null) continue;
			            System.out.println(getValue(one));
			            break;
			        }
			    }
			}
			//转换数据格式
		} catch (FileNotFoundException e) {
			e.printStackTrace();
		} catch (IOException e) {
			e.printStackTrace();
		}
	}
	private static String getValue(XSSFCell xssfRow) {
        if (xssfRow.getCellType() == XSSFCell.CELL_TYPE_BOOLEAN) {
            return String.valueOf(xssfRow.getBooleanCellValue());
        } else if (xssfRow.getCellType() == XSSFCell.CELL_TYPE_NUMERIC) {
            return String.valueOf(xssfRow.getNumericCellValue());
        } else {
            return String.valueOf(xssfRow.getStringCellValue());
        }
    }
	
}
