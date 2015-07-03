package com.adms.mglplanreport.service;

import java.io.IOException;
import java.util.ArrayList;
import java.util.Date;
import java.util.List;

import org.apache.commons.lang3.StringUtils;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.ss.util.CellRangeAddress;

import com.adms.mglplanlv.entity.ProductionByLot;
import com.adms.mglplanlv.service.productionbylot.ProductionByLotService;
import com.adms.mglplanreport.enums.ETemplateWB;
import com.adms.mglplanreport.util.ApplicationContextUtil;
import com.adms.mglplanreport.util.WorkbookUtil;
import com.adms.mglpplanreport.obj.MGLMonthlyObj;
import com.adms.utils.DateUtil;
import com.adms.utils.FileUtil;
import com.adms.utils.Logger;

public class MGLByCampaignMonthlyReport {
	
	private final int ALL_TEMPLATE_NUM = 3;
	
	private final int C_COLUMN_NUM = 2;
	private final int I_COLUMN_NUM = 8;
	
	private final int START_TABLE_HEADER_ROW = 7;
	private final int START_TABLE_DATA_ROW = 8;
	private final int TABLE_LIST_LOT_SUM_ROW = 9;
	private final int TABLE_MONTH_TO_DATE_SUM_ROW = 10;
	
	private final int TABLE_CAMPAIGN_TO_DATE_ROW = 9;
	
	private final String SHEET_NAME_PATTERN = "MMM-yy";
	private String EXPORT_FILE_NAME = "Production_Report-#campaignName_(forMGL)-#yyyyMMdd.xlsx";
	
	private MGLMonthlyObj sumProductionByMonth;
	
	private static Logger logger = Logger.getLogger();

	public void generateReport(String outPath, Date processDate) {
		logger.info("## MGL Production Report by Campaign ##");
		try {
//			Template
			Workbook wb = null;
			Sheet tempSheet = null;
			
//			receive All data
			List<MGLMonthlyObj> MGLObjs = getMGLDatas(DateUtil.convDateToString("yyyyMM", processDate));
			
			String campaignCode = "";
			String campaignName = "";
			String monthStr = "";
			logger.info("MGLObjs size: " + MGLObjs.size());
			for(MGLMonthlyObj mglMonthlyObj : MGLObjs) {
				
				if(!campaignCode.equals(mglMonthlyObj.getCampaignCode())) {
					
					if(StringUtils.isNotBlank(campaignCode)) {
//						sum all month
						tempSheet = wb.getSheetAt(ETemplateWB.MGL_BY_CAMPAIGN_TOTAL_TEMPLATE.getSheetIndex());
						doNewSheet(tempSheet, sumProductionByMonth, true);
						
//						do write out
						writeOut(wb, campaignName, processDate, outPath);
						monthStr = "";
					}
					
					campaignCode = mglMonthlyObj.getCampaignCode();
					campaignName = mglMonthlyObj.getCampaignName();

					sumProductionByMonth = new MGLMonthlyObj();
					sumProductionByMonth.setCampaignCode(campaignCode);
					sumProductionByMonth.setMonth("Total");
					
//					get template
					wb = WorkbookFactory.create(Thread.currentThread().getContextClassLoader().getResourceAsStream(ETemplateWB.MGL_BY_CAMPAIGN_TEMPLATE.getFilePath()));
					tempSheet = wb.getSheetAt(ETemplateWB.MGL_BY_CAMPAIGN_TEMPLATE.getSheetIndex());
					
				}
				
				if(!monthStr.equals(mglMonthlyObj.getMonth())) {
					try {
						doNewSheet(tempSheet, mglMonthlyObj, false);
						monthStr = mglMonthlyObj.getMonth();
					} catch(Exception e) {
						System.err.println("ERROR: " + mglMonthlyObj.toString());
						throw e;
					}
					
				}
				
			}
			
//			for last one
			if(StringUtils.isNotBlank(campaignCode)) {
//				sum all month
				doNewSheet(tempSheet, sumProductionByMonth, true);
				
//				do write out
				writeOut(wb, campaignName, processDate, outPath);
			}
			
		} catch (Exception e) {
			e.printStackTrace();
		}
	}
	
	private void doNewSheet(Sheet tempSheet, MGLMonthlyObj mglMonthlyObj, boolean isSumSheet) throws Exception {
//		new sheet
		Sheet toSheet = null;
		try {
			toSheet = tempSheet.getWorkbook().createSheet(mglMonthlyObj.getMonth());
		} catch(Exception e) {
			logger.error("Error when creating sheet name: " + mglMonthlyObj.getMonth());
			throw e;
		}
		
//		set Grid blank
		toSheet.setDisplayGridlines(false);
		
//		set report header
		setHeader(tempSheet, toSheet, mglMonthlyObj, isSumSheet);
		
//		set table
		if(isSumSheet) {
			doDataTableSumSheet(tempSheet, toSheet, mglMonthlyObj.getProductionByLots());
			doSumTableSumSheet(tempSheet, toSheet, mglMonthlyObj);
			
		} else {
			doDataTable(tempSheet, toSheet, mglMonthlyObj.getProductionByLots());
//			set sum table
			doSumTable(tempSheet, toSheet, mglMonthlyObj);
		}
	}
	
	private List<MGLMonthlyObj> getMGLDatas(String yearMonth) throws Exception {
		ProductionByLotService service = (ProductionByLotService) ApplicationContextUtil.getApplicationContext().getBean("productionByLotService");
		String hql = " from ProductionByLot d "
				+ " where 1 = 1"
				+ " and CONVERT(nvarchar(6), d.productionDate, 112) <= ? "
				+ " and CONVERT(nvarchar(4), d.productionDate, 112) = ? "
//				+ " and d.listLot.listLotCode in ('AAM15', 'AAN15') "
				+ " order by d.listLot.campaign.campaignCode, CONVERT(nvarchar(6), d.productionDate, 112), d.listLot.listLotCode ";
		List<ProductionByLot> prodList = service.findByHql(hql, yearMonth, yearMonth.substring(0, 4));
		
		if(prodList == null || prodList != null && prodList.size() == 0) {
			throw new Exception("Cannot get ProductionByLot: " + yearMonth + " | " + yearMonth.substring(0, 4));
		}
		
		List<MGLMonthlyObj> mglList = new ArrayList<MGLMonthlyObj>();
		MGLMonthlyObj mgl = null;
		
		String month = "";
		String campaignCode = "";
		
		for(ProductionByLot prod : prodList) {
			String currMonth = DateUtil.convDateToString(SHEET_NAME_PATTERN, prod.getProductionDate());
			String currCampaign = prod.getListLot().getCampaign().getCampaignCode();
			
			if(!campaignCode.equals(currCampaign) || !month.equals(currMonth)) {
				
				if(mgl != null) {
					mglList.add(mgl);
				}
				
				month = new String(currMonth);
				campaignCode = new String(currCampaign);
				
				mgl = new MGLMonthlyObj();
				mgl.setMonth(month);
				mgl.setCallingSite(prod.getListLot().getCampaign().getCallCenter().getCallCenterName());
				mgl.setCampaignCode(prod.getListLot().getCampaign().getCampaignCode());
				mgl.setCampaignName(prod.getListLot().getCampaign().getCampaignNameMgl());
				mgl.setRercordsReceived(prod.getTotalLead().doubleValue());
				
			}

			mgl.setCallingDate(DateUtil.convDateToString("yyyy-MM-dd", prod.getProductionDate()));
			if(StringUtils.isNotBlank(mgl.getListLot()) && !mgl.getListLot().contains(prod.getListLot().getListLotCode())) {
				mgl.setListLot(mgl.getListLot() + "," + prod.getListLot().getListLotCode());
			} else {
				mgl.setListLot(prod.getListLot().getListLotCode());
			}

			List<ProductionByLot> prods = mgl.getProductionByLots(); 
			if(mgl.getProductionByLots() == null) {
				prods = new ArrayList<>();
				mgl.setProductionByLots(prods);
			}
			mgl.getProductionByLots().add(prod);
			
		}
		mglList.add(mgl);
		
		return mglList;
	}
	
	private ProductionByLot getSumByListlotProduction(String listLot, String yearMonth) throws Exception {
		ProductionByLotService service = (ProductionByLotService) ApplicationContextUtil.getApplicationContext().getBean("productionByLotService");
		
		List<ProductionByLot> list = service.findByNamedQuery("findSumByListLotAndMonthProductionByLot", listLot, yearMonth);
		if(list == null || list != null && list.size() == 0) throw new Exception("Data not found: listLot: " + listLot + ", yearMonth: " + yearMonth);
		
		return list.get(0);
	}
	
	private ProductionByLot getSumByMonthProduction(String yearMonth, String campaignCode) throws Exception {
		ProductionByLotService service = (ProductionByLotService) ApplicationContextUtil.getApplicationContext().getBean("productionByLotService");
		List<ProductionByLot> list = service.findByNamedQuery("findSumByMonthProductionByLot", yearMonth, campaignCode);
		if(list == null || list != null && list.size() == 0) throw new Exception("Data not found: yearMonth: " + yearMonth + " | campaignCode: " + campaignCode);
		
		return list.get(0);
	}
	
	private ProductionByLot getSumByCampaignProductionByLot(String yearMonth, String campaignCode) throws Exception {
		ProductionByLotService service = (ProductionByLotService) ApplicationContextUtil.getApplicationContext().getBean("productionByLotService");
		List<ProductionByLot> list = service.findByNamedQuery("findSumMonthOfCampaignProductionByLot", campaignCode);
		if(list == null || list != null && list.size() == 0) throw new Exception("Data not found: campaignCode: " + campaignCode);
		
		return list.get(0);
	}
	
	private void setHeader(Sheet tempSheet, Sheet toSheet, MGLMonthlyObj mglByMonth, boolean isSumSheet) throws Exception {
		
		Cell tempCell = null;
		Cell toCell = null;
		
		for(int n = 0; n < START_TABLE_HEADER_ROW + 1; n++) {
			Row tempRow = tempSheet.getRow(n);
			Row toRow = toSheet.createRow(n);
			toRow.setHeightInPoints(tempRow.getHeightInPoints());
			for(int c = 0; c < I_COLUMN_NUM + 2; c++) {
				tempCell = tempRow.getCell(c, Row.CREATE_NULL_AS_BLANK);
				toCell = toRow.createCell(c, tempCell.getCellType());
				toCell.setCellStyle(tempCell.getCellStyle());
				switch(n) {
				case 0 : 
					if(c == 0) {
						toCell.setCellValue(tempCell.getStringCellValue());
						toSheet.addMergedRegion(new CellRangeAddress(n, n, 0, 1));
					}
					break;
				case 1 :
					if(c == 0) {
						String title = "";
						if(isSumSheet) {
							title = tempCell.getStringCellValue().replace("#yyyy", DateUtil.convDateToString("yyyy", mglByMonth.getProductionByLots().get(0).getProductionDate()));
						} else {
							title = tempCell.getStringCellValue()
									.replace("MMMM, yyyy", DateUtil.convDateToString("MMMM, yyyy", DateUtil.convStringToDate("yyyy-MM-dd", mglByMonth.getCallingDate())));
						}
						toCell.setCellValue(title);
						toSheet.addMergedRegion(new CellRangeAddress(n, n, 0, 4));
						toCell.setCellStyle(tempCell.getCellStyle());
					}
					break;
				case 2 :
					if(c == C_COLUMN_NUM - 2 || c == I_COLUMN_NUM - 2) {
						toCell.setCellValue(tempCell.getStringCellValue());
					} else if(c == C_COLUMN_NUM) {
						toCell.setCellValue(mglByMonth.getCampaignCode());
					} else if(c == I_COLUMN_NUM) {
						if(!isSumSheet) toCell.setCellValue(mglByMonth.getCallingDate());
					} else if(c == C_COLUMN_NUM - 1 || c == I_COLUMN_NUM - 1) {
						toSheet.addMergedRegion(new CellRangeAddress(n, n, c - 1, c));
					} else if(c == I_COLUMN_NUM - 3) {
						toSheet.addMergedRegion(new CellRangeAddress(n, n, C_COLUMN_NUM, C_COLUMN_NUM + 3));
					} else if(c == I_COLUMN_NUM + 1) {
						toSheet.addMergedRegion(new CellRangeAddress(n, n, I_COLUMN_NUM, I_COLUMN_NUM + 1));
					}
					break;
				case 3 : 
					if(c == C_COLUMN_NUM - 2 || c == I_COLUMN_NUM - 2) {
						toCell.setCellValue(tempCell.getStringCellValue());
					} else if(c == C_COLUMN_NUM) {
						if(!isSumSheet) toCell.setCellValue(mglByMonth.getListLot());
					} else if(c == I_COLUMN_NUM) {
						if(!isSumSheet) toCell.setCellValue(mglByMonth.getCallingSite());
					} else if(c == C_COLUMN_NUM - 1 || c == I_COLUMN_NUM - 1) {
						toSheet.addMergedRegion(new CellRangeAddress(n, n, c - 1, c));
					} else if(c == I_COLUMN_NUM - 3) {
						toSheet.addMergedRegion(new CellRangeAddress(n, n, C_COLUMN_NUM, C_COLUMN_NUM + 3));
					} else if(c == I_COLUMN_NUM + 1) {
						toSheet.addMergedRegion(new CellRangeAddress(n, n, I_COLUMN_NUM, I_COLUMN_NUM + 1));
					}
					break;
				case 4 : 
					if(c == C_COLUMN_NUM - 2) {
						toCell.setCellValue(tempCell.getStringCellValue());
					} else if(c == C_COLUMN_NUM) {
						if(!isSumSheet) toCell.setCellValue(mglByMonth.getRercordsReceived());
					} else if(c == C_COLUMN_NUM - 1 || c == I_COLUMN_NUM - 1) {
						toSheet.addMergedRegion(new CellRangeAddress(n, n, c - 1, c));
					} else if(c == I_COLUMN_NUM - 3) {
						toSheet.addMergedRegion(new CellRangeAddress(n, n, C_COLUMN_NUM, C_COLUMN_NUM + 3));
					}
					break;
				case 5 : 
					if(!isSumSheet) 
						if(c == C_COLUMN_NUM - 2) {
							toCell.setCellValue(tempCell.getStringCellValue());
						} else if(c == C_COLUMN_NUM) {
							toCell.setCellValue(DateUtil.convDateToString("dd/MM/yyyy h:mm:ss a", DateUtil.getCurrentDate()));
						} else if(c == C_COLUMN_NUM - 1 || c == I_COLUMN_NUM - 1) {
							toSheet.addMergedRegion(new CellRangeAddress(n, n, c - 1, c));
						} else if(c == I_COLUMN_NUM - 3) {
							toSheet.addMergedRegion(new CellRangeAddress(n, n, C_COLUMN_NUM, C_COLUMN_NUM + 3));
						}
					break;
				default : break;
				}
			}
			
		}
		
	}
	
	private void doDataTable(Sheet tempSheet, Sheet toSheet, List<ProductionByLot> productions) throws Exception {
//		Table header
		Row tempRow = tempSheet.getRow(START_TABLE_HEADER_ROW);
		Row toRow = toSheet.createRow(START_TABLE_HEADER_ROW);
		toRow.setHeightInPoints(tempRow.getHeightInPoints());
		
		for(int c = 0; c < tempRow.getLastCellNum(); c++) {
			Cell tempCell = tempRow.getCell(c, Row.CREATE_NULL_AS_BLANK);
			Cell toCell = toRow.createCell(c, tempCell.getCellType());
			WorkbookUtil.getInstance().copyCellValue(tempCell, toCell);
			WorkbookUtil.getInstance().copyColumnWidth(tempSheet, c, toSheet, c);
			toCell.setCellStyle(tempCell.getCellStyle());
		}
		
//		Table data
		tempRow = null;
		toRow = null;
		String listLot = "";
		String yearMonth = "";
		int currRow = START_TABLE_DATA_ROW;
		
		for(ProductionByLot product : productions) {
			
//			create row
			toRow = toSheet.createRow(currRow);
			
			if(!listLot.equals(product.getListLot().getListLotCode())) {
				if(StringUtils.isNotBlank(listLot)) {
//					sum by list lot row
					tempRow = tempSheet.getRow(TABLE_LIST_LOT_SUM_ROW);
					
//					do sum by list lot
					ProductionByLot sumByListlot = getSumByListlotProduction(listLot, yearMonth);
					setDataValue(tempRow, toRow, sumByListlot, true, false, false);
					currRow++;
					
//					create row
					toRow = toSheet.createRow(currRow);
					
				}

				listLot = product.getListLot().getListLotCode();
				yearMonth = DateUtil.convDateToString("yyyyMM", product.getProductionDate());
			}
			
//			data row
			tempRow = tempSheet.getRow(START_TABLE_DATA_ROW);
			setDataValue(tempRow, toRow, product, false, false, false);
			currRow++;
		}
//		create row
		toRow = toSheet.createRow(currRow);
//		sum by list lot row
		tempRow = tempSheet.getRow(TABLE_LIST_LOT_SUM_ROW);
//		do sum by list lot
		ProductionByLot sumByListlot = getSumByListlotProduction(listLot, yearMonth);
		setDataValue(tempRow, toRow, sumByListlot, true, false, false);
	}
	
	private void doDataTableSumSheet(Sheet tempSheet, Sheet toSheet, List<ProductionByLot> productions) throws Exception {
//		Table header
		Row tempRow = tempSheet.getRow(START_TABLE_HEADER_ROW);
		Row toRow = toSheet.createRow(START_TABLE_HEADER_ROW);
		toRow.setHeightInPoints(tempRow.getHeightInPoints());
		
		for(int c = 0; c < tempRow.getLastCellNum(); c++) {
			Cell tempCell = tempRow.getCell(c, Row.CREATE_NULL_AS_BLANK);
			Cell toCell = toRow.createCell(c, tempCell.getCellType());
			WorkbookUtil.getInstance().copyCellValue(tempCell, toCell);
			WorkbookUtil.getInstance().copyColumnWidth(tempSheet, c, toSheet, c);
			toCell.setCellStyle(tempCell.getCellStyle());
		}
		
//		Table data
		tempRow = null;
		toRow = null;
		int currRow = START_TABLE_DATA_ROW;
		
		tempRow = tempSheet.getRow(START_TABLE_DATA_ROW);
		
		for(ProductionByLot product : productions) {
//			create row
			toRow = toSheet.createRow(currRow++);
			setDataValue(tempRow, toRow, product, false, false, true);
		}
		
	}
	
	private void doSumTableSumSheet(Sheet tempSheet, Sheet toSheet, MGLMonthlyObj mglMonthlyObj) throws Exception {
		Row tempRow = tempSheet.getRow(TABLE_CAMPAIGN_TO_DATE_ROW);
		Row toRow = toSheet.createRow(toSheet.getLastRowNum() + 1);
		toRow.setHeightInPoints(tempRow.getHeightInPoints());
		
		ProductionByLot product = getSumByCampaignProductionByLot("", mglMonthlyObj.getCampaignCode() );
		setDataValue(tempRow, toRow, product, false, true, true);
	}
	
	private void doSumTable(Sheet tempSheet, Sheet toSheet, MGLMonthlyObj mglMonthlyObj) throws Exception {
		Row tempRow = tempSheet.getRow(TABLE_MONTH_TO_DATE_SUM_ROW);
		Row toRow = toSheet.createRow(toSheet.getLastRowNum() + 1);
		toRow.setHeightInPoints(tempRow.getHeightInPoints());
		
		ProductionByLot product = getSumByMonthProduction(DateUtil.convDateToString("yyyyMM", DateUtil.convStringToDate(SHEET_NAME_PATTERN, mglMonthlyObj.getMonth())), mglMonthlyObj.getCampaignCode() );
		setDataValue(tempRow, toRow, product, false, true, false);
		
		sumProductionByMonth.setCampaignCode(mglMonthlyObj.getCampaignCode());
		if(sumProductionByMonth.getProductionByLots() == null) {
			sumProductionByMonth.setProductionByLots(new ArrayList<ProductionByLot>());
		}
		sumProductionByMonth.getProductionByLots().add(product);
	}
	
	private void writeOut(Workbook wb, String campaignName, Date processDate, String outPath) throws IOException {
		
//		remove template sheet(s)
		for(int r = 0; r < ALL_TEMPLATE_NUM; r++) {
			wb.removeSheetAt(0);
		}
		
		String dateFormat = "yyyyMMdd";
		String outName = EXPORT_FILE_NAME.replaceAll("#campaignName", campaignName)
				.replaceAll("#".concat(dateFormat), DateUtil.convDateToString(dateFormat, processDate));
		
		FileUtil.getInstance().createDirectory(outPath);
		String outDir = WorkbookUtil.getInstance().writeOut(wb, outPath, outName);
		wb.close();
		wb = null;
		logger.info("Writed: " + campaignName + " | to: " + outDir);
	}
	
	private void setDataValue(Row tempRow, Row toRow, ProductionByLot product, boolean isSumListLot, boolean isSumMonth, boolean isSumSheet) throws Exception {
		Double hours = 0D;
		Double minutes = 0D;
		
		for(int c = 0; c < tempRow.getLastCellNum(); c++) {
			Cell tempCell = tempRow.getCell(c, Row.CREATE_NULL_AS_BLANK);
			Cell toCell = toRow.createCell(c, tempCell.getCellType());
			toCell.setCellStyle(tempCell.getCellStyle());
			switch(c) {
			case 0 : 
				/* <-- Day Column --> */
				toCell.setCellValue(isSumMonth ? tempCell.getStringCellValue() 
						: isSumListLot ? product.getListLot().getListLotCode() 
								: isSumSheet ? DateUtil.convDateToString(SHEET_NAME_PATTERN, product.getProductionDate()) 
										: DateUtil.getDayString(product.getProductionDate()));
				break;
			case 1 : 
				/* Date Column */
				if(!(isSumMonth || isSumListLot || isSumSheet)) toCell.setCellValue(DateUtil.convDateToString("yyyy-MM-dd", product.getProductionDate()));
				break;
			case 2 : 
				/* Hours */
				hours = product.getHour() + (product.getSecond() / 60D + product.getMinute()) / 60D;
				toCell.setCellValue(hours);
				break;
			case 3 : 
				/* Mnutes */
				minutes = product.getMinute() + product.getHour() * 60D + product.getSecond() / 60D;
				toCell.setCellValue(minutes);
				break;
			case 4 :
				/* HH:MM:SS */
				if(!(isSumMonth || isSumSheet)) toCell.setCellValue(DateUtil.getTotalHHmmss(product.getHour(), product.getMinute(), product.getSecond()));
				break;
			case 5 : 
				/* Dialing */
				toCell.setCellValue(product.getDialing().intValue());
				break;
			case 6 : 
				/* Completed */
				toCell.setCellValue(product.getCompleted().intValue());
				break;
			case 7 : 
				/* Con/Hour */
				if(!(isSumMonth || isSumSheet)) toCell.setCellValue(product.getHour() == 0L ? 0 : product.getCompleted() / product.getHour().doubleValue());
				break;
			case 8 : 
				/* Comp/Dial */
				if(!(isSumMonth || isSumSheet)) toCell.setCellValue(product.getDialing() == 0L ? 0 : product.getCompleted() / product.getDialing().doubleValue());
				break;
			case 9 : 
				/* Contact */
				toCell.setCellValue(product.getContact().intValue());
				break;
			case 10 : 
				/* Con/Hour */
				if(!(isSumMonth || isSumSheet)) toCell.setCellValue(product.getHour() == 0L ? 0 : product.getContact() / product.getHour().doubleValue());
				break;
			case 11 : 
				/* Con/Dial */
				if(!(isSumMonth || isSumSheet)) toCell.setCellValue(product.getDialing() == 0L ? 0 : product.getContact() / product.getDialing().doubleValue());
				break;
			case 12 : 
				/* Sales */
				toCell.setCellValue(product.getSales());
				break;
			case 13 : 
				/* SPH */
				if(!(isSumMonth || isSumSheet)) toCell.setCellValue(product.getHour() == 0L ? 0 : product.getSales() / product.getHour().doubleValue());
				break;
			case 14 : 
				/* SPCon% */
				if(!(isSumMonth || isSumSheet)) toCell.setCellValue(product.getContact() == 0L ? 0 : product.getSales() / product.getContact().doubleValue());
				break;
			case 15 : 
				/* ABAN% */
				if(!(isSumMonth || isSumSheet)) toCell.setCellValue(product.getAbandons());
				break;
			case 16 : 
				/* Abandons */
				if(!(isSumMonth || isSumSheet)) toCell.setCellValue(product.getAbandons());
				break;
			case 17 : 
				/* UW Replease Sales */
				if(!(isSumMonth || isSumSheet)) toCell.setCellValue(product.getUwReleaseSales());
				break;
			case 18 : 
				/* TYP */
				toCell.setCellValue(product.getTyp().doubleValue());
				break;
			case 19 : 
				/* TMP = TYP / 12 */ 
				toCell.setCellValue(product.getTyp().doubleValue() / 12D);
				break;
			case 20 : 
				/* AMP = TYP / 12 / Sales */
				toCell.setCellValue(product.getSales() == 0L ? 0 : product.getTyp().doubleValue() / 12D / product.getSales().doubleValue() );
				break;
			case 21 : 
				/* AYP = TYP / Sales */
				toCell.setCellValue(product.getSales() == 0L ? 0 : product.getTyp().doubleValue() / product.getSales().doubleValue());
				break;
			case 22 : 
				/* Total Cost */
				toCell.setCellValue(product.getTotalCost().doubleValue());
				break;
			case 24 : 
				/* Release Sales */
				if(!(isSumMonth || isSumSheet)) toCell.setCellValue(product.getReleaseSales());
				break;
			case 25 : 
				/* NET SPC% */
				if(!(isSumMonth || isSumSheet)) toCell.setCellValue(product.getContact() == 0L ? 0 : product.getSales() / product.getContact().doubleValue());
				break;
			case 26 : 
				/* TMP */
				if(!(isSumMonth || isSumSheet)) toCell.setCellValue(product.getTyp().doubleValue() / 12D);
				break;
			case 27 : 
				/* AMP Post UW */
				if(!(isSumMonth || isSumSheet)) toCell.setCellValue(product.getAmpPostUw().doubleValue());
				break;
			case 28 : 
				/* Declines */
				if(!(isSumMonth || isSumSheet)) toCell.setCellValue(product.getDeclines());
				break;
			case 29 : 
				/* Decline Rate% */
				if(!(isSumMonth || isSumSheet)) toCell.setCellValue(product.getContact() == 0L ? 0 : product.getDeclines() / product.getContact().doubleValue());
				break;
			default : break;
			}
		}
	}
}
