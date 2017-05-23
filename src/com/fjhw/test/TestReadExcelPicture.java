package com.fjhw.test;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.InputStream;
import java.util.HashMap;
import java.util.List;
import java.util.Map;
import java.util.UUID;

import org.apache.poi.hssf.usermodel.HSSFClientAnchor;
import org.apache.poi.hssf.usermodel.HSSFPatriarch;
import org.apache.poi.hssf.usermodel.HSSFPicture;
import org.apache.poi.hssf.usermodel.HSSFPictureData;
import org.apache.poi.hssf.usermodel.HSSFShape;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.PictureData;
import org.apache.poi.ss.usermodel.WorkbookFactory;

public class TestReadExcelPicture {
	public static void main(String[] args) throws Exception {
		InputStream inp = new FileInputStream("D:\\test.xls");
		HSSFWorkbook workbook = (HSSFWorkbook) WorkbookFactory.create(inp);
		List<HSSFPictureData> pictures = workbook.getAllPictures();
		// 获取图片所在的单表
		HSSFSheet sheet = (HSSFSheet) workbook.getSheetAt(1);
		Map<String, Object> map = new HashMap<String, Object>();
		// 若没有图片，则停止执行，并说明无图片
		HSSFPatriarch patriarch = sheet.getDrawingPatriarch();
		if (patriarch == null) {
			return;
		}
		List<HSSFShape> shapes = patriarch.getChildren();
		for (HSSFShape shape : shapes) {
			HSSFClientAnchor anchor = (HSSFClientAnchor) shape.getAnchor();
			if (shape instanceof HSSFPicture) {
				HSSFPicture pic = (HSSFPicture) shape;
				int row0 = anchor.getRow1();
				int col0 = anchor.getCol1();
				int row1 = anchor.getRow2();
				int col1 = anchor.getCol2();
				System.out.println("列:" + col0 + "-->" + col1);
				System.out.println("行:" + row0 + "-->" + row1);
				// map.put(row+":"+col, row+":"+col);
				int pictureIndex = pic.getPictureIndex() - 1;
				HSSFPictureData picData = pictures.get(pictureIndex);
				System.out.println("--->" + picData);
				map.put(row1 + ":" + col1, picData);
				savePic(UUID.randomUUID().toString(), picData);
			}
		}

		System.out.println(map);
	}

	private static void savePic(String i, PictureData pic) throws Exception {
		String ext = pic.suggestFileExtension();
		byte[] data = pic.getData();
		if (ext.equals("jpeg")) {
			FileOutputStream out = new FileOutputStream("E:\\" + i + ".jdp");
			out.write(data);
			out.close();
			File file = new File("E:\\" + i + ".jpg");
			FileInputStream in = new FileInputStream(file);
			System.out.println("in===>" + in);
			if (file.isFile()) {
				file.delete();
				System.out.println("=============delete");
			}
		} else {
			FileOutputStream out = new FileOutputStream("E:\\" + i + "." + ext);
			out.write(data);
			out.close();

		}

		/*
		 * if (ext.equals("png")) { FileOutputStream out = new
		 * FileOutputStream("F:\\" + i + ".jpg"); out.write(data); out.close();
		 * }
		 */
	}
}
