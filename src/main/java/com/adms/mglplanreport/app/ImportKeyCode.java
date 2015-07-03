package com.adms.mglplanreport.app;

import java.io.FileInputStream;
import java.io.InputStream;

import org.apache.commons.lang3.StringUtils;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;

import com.adms.mglplanlv.entity.TempKeyCodeInPast;
import com.adms.mglplanlv.service.mgltarget.TempKeyCodeInPastService;
import com.adms.mglplanreport.util.ApplicationContextUtil;

public class ImportKeyCode {

	public static void main(String[] args) {
		try {
			String path = "d:/keycode.xlsx";
			InputStream is = new FileInputStream(path);
			TempKeyCodeInPastService service = (TempKeyCodeInPastService) ApplicationContextUtil.getApplicationContext().getBean("tempKeyCodeInPastService");
			
			Workbook wb = WorkbookFactory.create(is);
			
			Sheet sheet = wb.getSheetAt(1);
			System.out.println(sheet.getLastRowNum());
			for(int n = 0; n < sheet.getLastRowNum(); n++) {
				Row row = sheet.getRow(n);
				
				if(row == null) continue;
				
				Cell cell = row.getCell(0, Row.CREATE_NULL_AS_BLANK);
				if(cell != null && StringUtils.isNoneBlank(cell.getStringCellValue())) {
					System.out.println(cell.getStringCellValue());
					service.add(new TempKeyCodeInPast(cell.getStringCellValue()), "Import System");
				}
			}
			
			System.out.println("finished");
			wb.close();
			is.close();
			
		} catch(Exception e) {
			e.printStackTrace();
		}
	}
}
