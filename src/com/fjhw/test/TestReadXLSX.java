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

/**����xlsx��ʽ��excel*/
public class TestReadXLSX {
	@Test
	public void test1(){
		try {
			InputStream is = new FileInputStream("���磺����������Ϣ����վ����ƽ̨���ݵ����.xlsx");
			XSSFWorkbook xssfWorkbook = new XSSFWorkbook(is);
			// ��ȡÿһ��������
			for (int numSheet = 0; numSheet < xssfWorkbook.getNumberOfSheets(); numSheet++) {
			    XSSFSheet xssfSheet = xssfWorkbook.getSheetAt(numSheet);
			    if (xssfSheet == null) {
			        continue;
			    }
			    // ��ȡ��ǰ��������ÿһ��
			    for (int rowNum = 1; rowNum <= xssfSheet.getLastRowNum(); rowNum++) {
			        XSSFRow xssfRow = xssfSheet.getRow(rowNum);
			        if (xssfRow != null) {
			        	//��ȡ��һ������
			            XSSFCell one = xssfRow.getCell(0);
			            if(one==null) continue;
			            System.out.println(getValue(one));
			            break;
			        }
			    }
			}
			//ת�����ݸ�ʽ
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
