package com.adms.mglplanreport.util;

import java.io.File;
import java.io.FileOutputStream;
import java.io.OutputStream;
import java.util.Date;

import org.apache.poi.hssf.usermodel.HSSFFormulaEvaluator;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.ClientAnchor;
import org.apache.poi.ss.usermodel.CreationHelper;
import org.apache.poi.ss.usermodel.Drawing;
import org.apache.poi.ss.usermodel.Picture;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;

import com.adms.utils.FileUtil;


public class WorkbookUtil extends org.apache.poi.ss.util.WorkbookUtil {

	private static WorkbookUtil instance;
	
	public static WorkbookUtil getInstance() {
		if(instance == null) {
			instance = new WorkbookUtil();
		}
		return instance;
	}
	
	public void addPicture(Sheet sheet, byte[] bytes, int rowIdx, int colIdx, int pictureType) {
		addPicture(sheet, bytes, rowIdx, colIdx, 1, pictureType);
	}
	
	public void addPicture(Sheet sheet, byte[] bytes, int rowIdx, int colIdx, int resizePercent, int pictureType) {
		int pictureIdx = sheet.getWorkbook().addPicture(bytes, pictureType);
		
		CreationHelper helper = sheet.getWorkbook().getCreationHelper();
		Drawing drawing = sheet.createDrawingPatriarch();
		
		ClientAnchor anchor = helper.createClientAnchor();
		anchor.setRow1(rowIdx);
		anchor.setCol1(colIdx);
		
		Picture picture = drawing.createPicture(anchor, pictureIdx);
		picture.resize(resizePercent);
	}
	
	public void refreshAllFormula(Workbook wb) {
		HSSFFormulaEvaluator.evaluateAllFormulaCells(wb);
	}
	
	public void copyCellValue(Cell origCell, Cell toCell) {
		switch(origCell.getCellType()) {
		case Cell.CELL_TYPE_BLANK :
			toCell.setCellValue(origCell.getStringCellValue());
			break;
		case Cell.CELL_TYPE_BOOLEAN :
			toCell.setCellValue(origCell.getBooleanCellValue());
			break;
		case Cell.CELL_TYPE_ERROR :
			toCell.setCellValue(origCell.getErrorCellValue());
			break;
		case Cell.CELL_TYPE_FORMULA :
			toCell.setCellValue(origCell.getCellFormula());
			break;
		case Cell.CELL_TYPE_NUMERIC :
			toCell.setCellValue(origCell.getNumericCellValue());
			break;
		case Cell.CELL_TYPE_STRING :
			toCell.setCellValue(origCell.getRichStringCellValue());
			break;
		}
	}
	
	public void setObjectValueToCell(Cell cell, Object value) throws Exception {
		if(value instanceof String) {
			cell.setCellValue(String.valueOf(value));
		} else if(value instanceof Integer) {
			cell.setCellValue(Integer.valueOf(String.valueOf(value)));
		} else if(value instanceof Date) {
			cell.setCellValue((Date) value);
		} else if(value instanceof Double) {
			cell.setCellValue(Double.valueOf(String.valueOf(value)));
		} else {
			throw new Exception("ERROR: Instance not found for object value: " + value.getClass().getName());
		}
	}
	
	public void copyColumnWidth(Sheet origSheet, int origColumnNum, Sheet toSheet, int toColumnNum) {
		toSheet.setColumnWidth(toColumnNum, origSheet.getColumnWidth(origColumnNum));
	}
	
	public String writeOut(Workbook wb, String dir, String fileName) {
		try {
			String outFileDir = "";
			FileUtil.getInstance().createDirectory(dir);
			outFileDir = dir + File.separatorChar + fileName;
			OutputStream os = new FileOutputStream(outFileDir);
			wb.write(os);
			os.close();
			return outFileDir;
		} catch(Exception e) {
			e.printStackTrace();
		}
		return null;
	}
}
