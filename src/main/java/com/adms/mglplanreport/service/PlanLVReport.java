package com.adms.mglplanreport.service;

import java.io.IOException;
import java.util.Date;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;

import com.adms.mglplanlv.entity.Campaign;
import com.adms.mglplanlv.service.campaign.CampaignService;
import com.adms.mglplanreport.enums.ETemplateWB;
import com.adms.mglplanreport.service.planlv.PlanLevelGenerator;
import com.adms.mglplanreport.service.planlv.factory.PlanLevelGeneratorFactory;
import com.adms.mglplanreport.util.ApplicationContextUtil;
import com.adms.mglplanreport.util.WorkbookUtil;
import com.adms.mglpplanreport.obj.PlanLevelObj;
import com.adms.utils.DateUtil;
import com.adms.utils.FileUtil;
import com.adms.utils.Logger;

public class PlanLVReport {
	
	private final String EXPORT_FILE_NAME = "Production-PlanLV-YTD-#MMM_yyyyMM.xlsx";
	private int _all_template_num = 0;
	
	private Map<String, Integer> _campaignSheetIdxMap;
	
	private static Logger logger = Logger.getLogger();
	
	public PlanLVReport() {
		try {
			logger.info("## Plan Level Report ##");
			initCampaignSheet(WorkbookFactory.create(ClassLoader.getSystemResourceAsStream(ETemplateWB.PLAN_LV_TEMPLATE.getFilePath())));
		} catch(Exception e) {
			logger.error(e.getMessage(), e);
		}
	}
	
	public void generateReport(String outPath, Date processDate) {
		try {

			Workbook wb = null;
			int sheetIdx = 999;
			
			List<Campaign> campaigns = getAllCampaignInYear(processDate);
			
			for(Campaign campaign : campaigns) {
				
//				<!-- Skip -->
//				if(campaign.getCampaignCode().equals("021PA1715M04")) continue;
				
				if(wb == null) {
					wb = WorkbookFactory.create(ClassLoader.getSystemResourceAsStream(ETemplateWB.PLAN_LV_TEMPLATE.getFilePath()));
					_all_template_num = wb.getNumberOfSheets();
				}
				
				try {
					PlanLevelGenerator planLv = PlanLevelGeneratorFactory.getGenerator(campaign.getCampaignNameMgl());
					logger.info("Getting MTD Data: " + campaign.getCampaignCode() + " | processDate: " + DateUtil.convDateToString("yyyyMMdd", processDate));
					PlanLevelObj mtdData = planLv.getMTDData(campaign.getCampaignCode(), processDate);
					
					logger.info("Getting YTD Data: " + campaign.getCampaignCode() + " | processDate: " + DateUtil.convDateToString("yyyyMMdd", processDate));
					PlanLevelObj ytdData = planLv.getYTDData(campaign.getCampaignCode(), processDate);
					
					sheetIdx = _campaignSheetIdxMap.get(campaign.getCampaignCode());
					
					if(sheetIdx == 999) {
						throw new Exception("Cannot find template sheet index: " + campaign.getCampaignCode());
					}
					
					planLv.generateDataSheet(wb.getSheetAt(sheetIdx), mtdData, ytdData);
					planLv = null;
				} catch(Exception e) {
					logger.error(e.getMessage(), e);
				}
			}
			
			writeOut(wb, processDate, outPath);
			
		} catch(Exception e) {
			logger.error(e.getMessage(), e);
		}
	}
	
	private void initCampaignSheet(Workbook wb) throws IOException {
		_campaignSheetIdxMap = new HashMap<>();
		System.out.println("wb.getNumberOfSheets(): " + wb.getNumberOfSheets());
		for(int i = 0; i < wb.getNumberOfSheets(); i++) {
			_campaignSheetIdxMap.put(wb.getSheetAt(i).getRow(0).getCell(1, Row.CREATE_NULL_AS_BLANK).getStringCellValue(), i);
		}
		wb.close();
	}

	private List<Campaign> getAllCampaignInYear(Date processDate) throws Exception {
		CampaignService service = (CampaignService) ApplicationContextUtil.getApplicationContext().getBean("campaignService");
		return service.findCampaignByLikeListLot("%"  + DateUtil.convDateToString("yy", processDate));
	}
	
	private void writeOut(Workbook wb, Date processDate, String outPath) throws IOException {
		
//		remove template sheet(s)
		for(int r = 0; r < _all_template_num; r++) {
			wb.removeSheetAt(0);
		}
		
//		Sorting Sheets
		sortingSheets(wb);
		
		for(int i = 0; i < wb.getNumberOfSheets(); i++) {
			wb.setSheetName(i, wb.getSheetAt(i).getSheetName().replace("(2)", "").trim());
		}
		
		String outName = EXPORT_FILE_NAME.replaceAll("#".concat("MMM_yyyyMM"), DateUtil.convDateToString("MMM_yyyyMM", processDate));
		
		FileUtil.getInstance().createDirectory(outPath);
		String outDir = WorkbookUtil.getInstance().writeOut(wb, outPath, outName);
		wb.close();
		wb = null;
		logger.info("Writed: " + outDir);
	}
	
	private void sortingSheets(Workbook wb) {
//		sorting sheets
		int len = wb.getNumberOfSheets();
		int k;
		System.out.println("_campaignSheetIdxMap size: " + _campaignSheetIdxMap.size());
		for(int n = len; n >= 0; n--) {
			for(int i = 0; i < len - 1; i++) {
				k = i + 1;
				String a = wb.getSheetAt(i).getRow(0).getCell(1).getStringCellValue();
				String b = wb.getSheetAt(k).getRow(0).getCell(1).getStringCellValue();
				
				if(_campaignSheetIdxMap.get(a) > _campaignSheetIdxMap.get(b)) {
					try {
						swap(wb, _campaignSheetIdxMap.get(a), _campaignSheetIdxMap.get(b));
					} catch(Exception e) {
						System.err.println("err: " + a + " cannot swap with " + b);
						logger.error(e.getMessage(), e);
					}
				}
			}
		}
		
	}
	
	private void swap(Workbook wb, int idxA, int idxB) throws Exception {
		try {
//			System.out.println("swap sheet length: " + wb.getNumberOfSheets());
			wb.setSheetOrder(wb.getSheetAt(idxA).getSheetName(), idxB);
		} catch(Exception e) {
			logger.error("Workbook sheet quantity: " + wb.getNumberOfSheets() + ", Cannot swap A: " + idxA + " and B: " + idxB);
			throw e;
		}
	}
}
