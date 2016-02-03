package com.adms.mglplanreport.service;

import java.io.IOException;
import java.util.Date;
import java.util.HashMap;
import java.util.List;
import java.util.Map;
import java.util.stream.Collectors;

import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.hibernate.criterion.DetachedCriteria;
import org.hibernate.criterion.Projections;
import org.hibernate.criterion.Restrictions;
import org.hibernate.type.StringType;

import com.adms.mglplanlv.entity.Campaign;
import com.adms.mglplanlv.entity.Sales;
import com.adms.mglplanlv.service.campaign.CampaignService;
import com.adms.mglplanlv.service.sales.SalesService;
import com.adms.mglplanreport.enums.ETemplateWB;
import com.adms.mglplanreport.service.planlv.PlanLevelGenerator;
import com.adms.mglplanreport.service.planlv.impl.AllPlanLvGenImpl;
import com.adms.mglplanreport.util.ApplicationContextUtil;
import com.adms.mglplanreport.util.WorkbookUtil;
import com.adms.mglpplanreport.obj.PlanLevelObj;
import com.adms.utils.DateUtil;
import com.adms.utils.FileUtil;
import com.adms.utils.Logger;

public class PlanLVReport {
	
	private final String EXPORT_FILE_NAME = "Production-PlanLV-YTD-#MMM_yyyyMM.xlsx";
	private int _all_template_num = 0;

	private Map<String, String> campaignCodeByName;
	private Map<String, Integer> _campaignSheetIdxMap;
	
	private static Logger logger = Logger.getLogger();
	
	public PlanLVReport() {
		try {
			ApplicationContextUtil.getApplicationContext();
			logger.info("## Plan Level Report ##");
			initCampaignSheet(WorkbookFactory.create(ClassLoader.getSystemResourceAsStream(ETemplateWB.PLAN_LV_TEMPLATE.getFilePath())));
		} catch(Exception e) {
			logger.error(e.getMessage(), e);
		}
	}
	
	private void initCampaignCode(String campaignYear) throws Exception {
		if(campaignCodeByName == null) {
			campaignCodeByName = new HashMap<>();
			CampaignService campaignService = (CampaignService) ApplicationContextUtil.getApplicationContext().getBean("campaignService");
			DetachedCriteria criteria = DetachedCriteria.forClass(Campaign.class);
			criteria.add(Restrictions.eq("campaignYear", campaignYear));
			List<Campaign> campaigns = campaignService.findByCriteria(criteria);
			campaignCodeByName = campaigns.stream()
					.collect(Collectors.toMap(Campaign::getCampaignNameMgl, Campaign::getCampaignCode));
		}
	}
	
	public void generateReport(String outPath, Date processDate) {
		try {

			Workbook wb = null;
			int sheetIdx;
			
			List<String> campaignNameList = getAllCampaignInYear(processDate);
			
			for(String campaignName : campaignNameList) {
				sheetIdx = -1;
				if(wb == null) {
					wb = WorkbookFactory.create(ClassLoader.getSystemResourceAsStream(ETemplateWB.PLAN_LV_TEMPLATE.getFilePath()));
					_all_template_num = wb.getNumberOfSheets();
				}
				
				try {
					PlanLevelGenerator planLv = new AllPlanLvGenImpl();
					
//					if(planLv == null) {
//						logger.warn("Plan Level Generator for '" + campaign.getCampaignNameMgl() + "' not found.");
//						continue;
//					}
					
					logger.info("Getting MTD Data: " + campaignName + " | processDate: " + DateUtil.convDateToString("yyyyMMdd", processDate));
					PlanLevelObj mtdData = null;
					
					try {
						mtdData = planLv.getMTDData(this.campaignCodeByName.get(campaignName), campaignName, processDate);
					} catch(Exception e) {
						logger.error("Cannot get MTD data for: " + campaignName + " | processDate: " + processDate);
						throw e;
					}
					
					logger.info("Getting YTD Data: " + campaignName + " | processDate: " + DateUtil.convDateToString("yyyyMMdd", processDate));
					PlanLevelObj ytdData = null;

					try {
						ytdData = planLv.getYTDData(this.campaignCodeByName.get(campaignName), campaignName, processDate);
					} catch(Exception e) {
						logger.error("Cannot get YTD data for: " + campaignName + " | processDate: " + processDate);
						throw e;
					}
					
					sheetIdx = -1;
					
					try {
						sheetIdx = _campaignSheetIdxMap.get(this.campaignCodeByName.get(campaignName));
					} catch(Exception e) {
						logger.error("Cannot find sheet template index for: '" + campaignName + "', campaignCode: " + this.campaignCodeByName.get(campaignName));
						logger.error("Exit...");
						System.exit(1);
						throw e;
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
		for(int i = 0; i < wb.getNumberOfSheets(); i++) {
			_campaignSheetIdxMap.put(wb.getSheetAt(i).getRow(0).getCell(1, Row.CREATE_NULL_AS_BLANK).getStringCellValue(), i);
		}
		wb.close();
	}

	@SuppressWarnings("unchecked")
	private List<String> getAllCampaignInYear(Date processDate) throws Exception {
		SalesService salesService = (SalesService) ApplicationContextUtil.getApplicationContext().getBean("salesService");
		
		String processYear = DateUtil.convDateToString("yyyy", processDate);
		initCampaignCode(processYear);
		
		DetachedCriteria sales = DetachedCriteria.forClass(Sales.class);
		sales.add(Restrictions.sqlRestriction("CONVERT(nvarchar(4), {alias}.SALE_DATE, 112) = ?", processYear, StringType.INSTANCE));
		
		DetachedCriteria campaign = sales.createCriteria("listLot", "l").createCriteria("campaign", "c");
		campaign.setProjection(Projections.distinct(Projections.property("c.campaignNameMgl")));
		
		List<?> list = salesService.findByCriteria(sales);
		return (List<String>) list.stream().collect(Collectors.toList());
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
		int len = wb.getNumberOfSheets();
		int k;

		for(int n = len; n >= 0; n--) {
			for(int i = 0; i < len - 1; i++) {
				k = i + 1;
				String a = wb.getSheetAt(i).getRow(0).getCell(1).getStringCellValue();
				String b = wb.getSheetAt(k).getRow(0).getCell(1).getStringCellValue();
				
				if(_campaignSheetIdxMap.get(a) > _campaignSheetIdxMap.get(b)) {
					try {
						swap(wb, i, i+1);
					} catch(Exception e) {
						logger.error(e.getMessage(), e);
					}
				}
			}
		}
	}
	
	private void swap(Workbook wb, int idxA, int idxB) throws Exception {
		try {
			wb.setSheetOrder(wb.getSheetAt(idxA).getSheetName(), idxB);
		} catch(Exception e) {
			logger.error("Workbook sheet quantity: " + wb.getNumberOfSheets() + ", Cannot swap A: " + idxA + " and B: " + idxB);
			throw e;
		}
	}
}
