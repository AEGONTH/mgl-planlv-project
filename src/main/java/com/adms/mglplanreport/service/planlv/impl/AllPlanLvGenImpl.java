package com.adms.mglplanreport.service.planlv.impl;

import java.util.Date;
import java.util.List;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;

import com.adms.mglplanlv.entity.PlanLvValue;
import com.adms.mglplanlv.service.planlv.PlanLvValueService;
import com.adms.mglplanreport.service.planlv.AbstractPlanLevelGenerator;
import com.adms.mglplanreport.util.ApplicationContextUtil;
import com.adms.mglplanreport.util.WorkbookUtil;
import com.adms.mglpplanreport.obj.PlanLevelObj;
import com.adms.utils.DateUtil;
import com.adms.utils.Logger;

public class AllPlanLvGenImpl extends AbstractPlanLevelGenerator {

	private static Logger logger = Logger.getLogger();

	@Override
	public PlanLevelObj getMTDData(String campaignCode, String campaignName, Date processDate) throws Exception {
		PlanLvValueService service = (PlanLvValueService) ApplicationContextUtil.getApplicationContext().getBean("planLvValueService");
//		[0] is campaignCode, [1] is approve yearMonth
		List<PlanLvValue> planLvList = service.findByNamedQuery("execPlanLvAll", "MTD", campaignName, DateUtil.convDateToString("yyyyMM", processDate));
		PlanLevelObj result = new PlanLevelObj();
		result.setCampaignCode(campaignCode);
		result.setMonthYear(DateUtil.convDateToString("MMM-yy", processDate));
		result.setPlanLvValues(planLvList);
		return result;
	}

	@Override
	public PlanLevelObj getYTDData(String campaignCode, String campaignName, Date processDate) throws Exception {
		PlanLvValueService service = (PlanLvValueService) ApplicationContextUtil.getApplicationContext().getBean("planLvValueService");
//		[0] is campaignCode, [1] is approve yearMonth
		List<PlanLvValue> planLvList = service.findByNamedQuery("execPlanLvAll", "YTD", campaignName, DateUtil.convDateToString("yyyyMM", processDate));
		PlanLevelObj result = new PlanLevelObj();
		result.setCampaignCode(campaignCode);
		result.setMonthYear(DateUtil.convDateToString("yyyy", processDate));
		result.setPlanLvValues(planLvList);
		return result;
	}

	@Override
	public void generateDataSheet(Sheet tempSheet, PlanLevelObj planLevelMtdObj, PlanLevelObj planLevelYTDObj) throws Exception {
		Workbook wb = tempSheet.getWorkbook();
		
		Sheet sheet = wb.cloneSheet(wb.getSheetIndex(tempSheet));
		
		Cell cell = sheet.getRow(2).getCell(0, Row.CREATE_NULL_AS_BLANK);
		cell.setCellValue(planLevelMtdObj.getMonthYear());

		if(planLevelMtdObj.getPlanLvValues() != null && planLevelMtdObj.getPlanLvValues().size() > 0) {
			setDataToTable(sheet, planLevelMtdObj.getPlanLvValues(), "MTD");
		} else {
			logger.warn("MTD Data is null: " + planLevelMtdObj.getCampaignCode());
		}
		
		if(planLevelYTDObj.getPlanLvValues() != null && planLevelYTDObj.getPlanLvValues().size() > 0) {
			setDataToTable(sheet, planLevelYTDObj.getPlanLvValues(), "YTD");
		} else {
			logger.warn("YTD Data is null: " + planLevelYTDObj.getCampaignCode());
		}
		sheet.setPrintGridlines(false);
	}
	
	private void setDataToTable(Sheet sheet, List<PlanLvValue> planLvList, String section) throws Exception {
//		Individual are Cols B to E
//		Spouse are Cols F to I
		int planIdx = 0;
		boolean isMtd = section.toUpperCase().equals("MTD");
		
		int noOfFileRowIdx = isMtd ? 3 : 7;
		int typRowIdx = isMtd ? 4 : 8;
//		int ampRowIdx = isMtd ? 5 : 9;
		
		for(PlanLvValue planLv : planLvList) {
			String planType = "";
			try {
				planType = planLv.getPlanType().toUpperCase();
			} catch(Exception e) {
				logger.error("Error Plan Type: " + planLv.toString());
				throw e;
			}
			
			try {
				planIdx = getPlanColumnIdx(sheet, planLv.getProduct().toUpperCase(), planType);
			} catch(Exception e) {
				logger.error("Error while get Plan Column Index..: " 
						+ sheet.getSheetName() 
						+ " | product: " + planLv.getProduct()
						+ " | planType: " + planType);
				throw e;
			}
			
			if(planIdx == -1) {
				logger.error("Not found Plan Column Index..: " 
						+ sheet.getSheetName() 
						+ " | product: " + planLv.getProduct()
						+ " | planType: " + planType);
			}
			
			sheet.getRow(noOfFileRowIdx).getCell(planIdx).setCellValue(planLv.getNumOfFile());
			sheet.getRow(typRowIdx).getCell(planIdx).setCellValue(planLv.getTyp());
//			sheet.getRow(ampRowIdx).getCell(planIdx).setCellValue(planLv.getAmp());
		}

		WorkbookUtil.getInstance().refreshAllFormula(sheet.getWorkbook());
	}

	private int getPlanColumnIdx(Sheet sheet, String productType, String planType) {
		int planRowIdx = 2;
		int beginCol = -1;
		
		for(int n = 0; n < sheet.getRow(planRowIdx - 1).getLastCellNum(); n++) {
			Cell cell = sheet.getRow(planRowIdx - 1).getCell(n, Row.CREATE_NULL_AS_BLANK);
			beginCol = retrieveBeginCol(cell, productType);
			if(beginCol != -1) break;
		}
		
		for(int i = beginCol; i < sheet.getRow(planRowIdx).getLastCellNum(); i++) {
			Cell cell = sheet.getRow(planRowIdx).getCell(i, Row.CREATE_NULL_AS_BLANK);
			if(cell.getStringCellValue().toUpperCase().contains(planType.toUpperCase())) {
				return i;
			}
		}
		return beginCol;
	}
	
	private int retrieveBeginCol(Cell cell, String productSection) {
		if(cell.getStringCellValue().toUpperCase().contains(productSection.toUpperCase())) {
			return cell.getColumnIndex();
		}
		return -1;
	}

}
