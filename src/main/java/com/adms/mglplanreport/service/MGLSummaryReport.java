package com.adms.mglplanreport.service;

import java.io.IOException;
import java.util.ArrayList;
import java.util.Date;
import java.util.HashMap;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Map;

import org.apache.commons.lang3.StringUtils;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.ss.util.CellReference;
import org.apache.poi.util.IOUtils;
import org.hibernate.criterion.DetachedCriteria;
import org.hibernate.criterion.Order;
import org.hibernate.criterion.Restrictions;
import org.hibernate.type.StringType;

import com.adms.mglplanlv.entity.Campaign;
import com.adms.mglplanlv.entity.ProductionByLot;
import com.adms.mglplanlv.service.campaign.CampaignService;
import com.adms.mglplanlv.service.productionbylot.ProductionByLotService;
import com.adms.mglplanreport.enums.ETemplateWB;
import com.adms.mglplanreport.util.ApplicationContextUtil;
import com.adms.mglplanreport.util.WorkbookUtil;
import com.adms.mglpplanreport.obj.MGLSummaryObj;
import com.adms.utils.DateUtil;
import com.adms.utils.FileUtil;
import com.adms.utils.Logger;

public class MGLSummaryReport {
	
//	private final String MTD_STR = "MTD";
//	private final String YTD_STR = "YTD";
	
	private int _all_template_num = 0;
//	private final int START_MTD_COL = 2;
	
	private final int START_TABLE_HEADER_ROW = 7;
	private final int START_TABLE_DATA_ROW = 9;
	private final int TEMP_TABLE_TOTAL_ROW = 12;
	private final String MONTH_PATTERN = DateUtil.getDefaultMonthPattern();
	
//	private Map<String, Double[]> sumOfMtdMap = new HashMap<>();
//	private Double[] sumAllOfYTD = new Double[]{0D, 0D};
	
//	private Map<String, Double> sumOfIAP = new HashMap<>();
	
	private final String EXPORT_FILE_NAME = "MGL_Summary_#yyyy.xlsx";
	
	private List<Integer> hideCols = new ArrayList<>();
	
	private static Logger logger = Logger.getLogger();
	
	public void generateReport(String outPath, Date processDate) {
		ApplicationContextUtil.getApplicationContext();
		logger.info("## Start MGL Summary Report ##");
		try {
			//Template
			Workbook wb = WorkbookFactory.create(Thread.currentThread().getContextClassLoader().getResourceAsStream(ETemplateWB.MGL_SUMMARY_TEMPLATE.getFilePath()));
			_all_template_num = wb.getNumberOfSheets();
			Sheet tempSheet = wb.getSheetAt(ETemplateWB.MGL_SUMMARY_TEMPLATE.getSheetIndex());
			Sheet toSheet = wb.createSheet("MGL_SUM");
			
//			set Grid blank
			toSheet.setDisplayGridlines(false);
			
//			set caption
			Cell captionCell = toSheet.createRow(5).createCell(0, tempSheet.getRow(5).getCell(0, Row.CREATE_NULL_AS_BLANK).getCellType());
			captionCell.setCellStyle(tempSheet.getRow(5).getCell(0, Row.CREATE_NULL_AS_BLANK).getCellStyle());
			String caption = tempSheet.getRow(5).getCell(0, Row.CREATE_NULL_AS_BLANK).getStringCellValue();
			String captionDateFormat = "MMM yyyy";
			captionCell.setCellValue(caption.replace(captionDateFormat, DateUtil.convDateToString(captionDateFormat, processDate)));
			
//			retrieving data
			List<MGLSummaryObj> mglSumList = getMGLSummary(processDate);
			int maxMonth = getMaxMonthInYear(mglSumList);
			
//			set table column header
			doTableHeader(tempSheet, toSheet, maxMonth);
			
//			set table data
			doTableData(tempSheet, toSheet, maxMonth, mglSumList, processDate);
			
//			set table total
			doTableTotal(tempSheet, toSheet, maxMonth);
			
//			hide cols
			if(hideCols.size() > 2) {
				int i = 0;
				for(int n = 0; n < hideCols.size() - 2; n++) {
					toSheet.setColumnWidth(hideCols.get(i++), 0);
				}
				
			}
			
//			insert picture
			byte[] bytes = IOUtils.toByteArray(Thread.currentThread().getContextClassLoader().getResourceAsStream("template/ADAMS_logo_th.png"));
			WorkbookUtil.getInstance().addPicture(toSheet, bytes, 0, 0, Workbook.PICTURE_TYPE_PNG);
			
//			remove template sheet(s)
			for(int r = 0; r < _all_template_num; r++) {
				wb.removeSheetAt(0);
			}

//			write out
			writeOut(wb, processDate, outPath);
			
		} catch (Exception e) {
			e.printStackTrace();
		} finally {
			
		}
		
	}
	
	private String getCampaignCode(String campaignYear, String campaignNameMgl) throws Exception {
		CampaignService campaignService = (CampaignService) ApplicationContextUtil.getApplicationContext().getBean("campaignService");
		DetachedCriteria criteria = DetachedCriteria.forClass(Campaign.class);
		criteria.add(Restrictions.eq("campaignYear", campaignYear))
			.add(Restrictions.eq("campaignNameMgl", campaignNameMgl));
		List<Campaign> campaigns = campaignService.findByCriteria(criteria);
		if(campaigns != null && campaigns.size() == 1) {
			return campaigns.get(0).getCampaignCode();
		} else {
			return null;
		}
	}

	private List<MGLSummaryObj> getMGLSummary(Date dataDate) throws Exception {
//		for Test
		logger.info("Get Production By Lot datas... by year to date");
		ProductionByLotService productionService = (ProductionByLotService) ApplicationContextUtil.getApplicationContext().getBean("productionByLotService");
		
		List<MGLSummaryObj> mglSumList = new ArrayList<>();
		
		Map<String, Double[]> mtdMap = null;
		MGLSummaryObj obj = null;
		
		DetachedCriteria criteria = DetachedCriteria.forClass(ProductionByLot.class);
		criteria.add(Restrictions.sqlRestriction("CONVERT(nvarchar(6), {alias}.PRODUCTION_DATE, 112) <= ?", DateUtil.convDateToString("yyyyMM", dataDate), StringType.INSTANCE));
		criteria.add(Restrictions.sqlRestriction("CONVERT(nvarchar(4), {alias}.PRODUCTION_DATE, 112) = ?", DateUtil.convDateToString("yyyy", dataDate), StringType.INSTANCE));
		DetachedCriteria listLot = criteria.createCriteria("listLot");
		DetachedCriteria campaign = listLot.createCriteria("campaign");
		campaign.addOrder(Order.asc("campaignNameMgl"));
		listLot.addOrder(Order.asc("listLotCode"));
		criteria.addOrder(Order.asc("productionDate"));
		
		List<ProductionByLot> productions = productionService.findByCriteria(criteria);
		
//		String hql = " from ProductionByLot d "
//				+ " where 1 = 1 "
//				+ " and CONVERT(nvarchar(6), d.productionDate, 112) <= ? "
//				+ " and CONVERT(nvarchar(4), d.productionDate, 112) = ? "
//				+ " order by d.listLot.campaign.campaignNameMgl, d.listLot.listLotCode, d.productionDate ";
//		List<ProductionByLot> productions = productionService.findByHql(hql, DateUtil.convDateToString("yyyyMM", dataDate), DateUtil.convDateToString("yyyy", dataDate));
		
		logger.info("Production data total records: " + productions.size());
		
		String campaignName = "";
		logger.info("Do Summarize...");
		for(ProductionByLot prod : productions) {
			
			if(!campaignName.equals(prod.getListLot().getCampaign().getCampaignNameMgl())){
				logger.info("From " + campaignName + " | to " + prod.getListLot().getCampaign().getCampaignNameMgl());
				
				if(StringUtils.isNoneBlank(campaignName)) {
					mglSumList.add(obj);
				}
				
				campaignName = prod.getListLot().getCampaign().getCampaignNameMgl();
				obj = new MGLSummaryObj();
				obj.setCampaignCode(getCampaignCode(DateUtil.convDateToString("yyyy", dataDate), campaignName));
				obj.setCampaignName(campaignName);

//				obj.setIssuedRate(mglTargetByCampaign.get(campaignCode).getIssuedRate().doubleValue());
//				obj.setPaidRate(mglTargetByCampaign.get(campaignCode).getPaidRate().doubleValue());

				obj.setIapMTD(0D);
				obj.setIapYTD(0D);
				mtdMap = new HashMap<>();
			}
			String mmm = DateUtil.convDateToString("MMM", prod.getProductionDate());
			
			Double[] mtds = mtdMap.get(mmm);
			if(mtds == null) {
				mtds = new Double[]{0D, 0D};
			}
			
			if(prod.getSales().doubleValue() > 0) {
				mtds[0] = mtds[0] + prod.getSales().doubleValue();
			}
			
			if(mtds[1] + prod.getTyp().doubleValue() > 0) {			
				mtds[1] = mtds[1] + prod.getTyp().doubleValue();
			}
			
			mtdMap.put(mmm, mtds);

			obj.setMTD(mtdMap);
		}
		mglSumList.add(obj);
		
		return mglSumList;
	}
	
	private int getMaxMonthInYear(List<MGLSummaryObj> mglSumList) {
		int max = 0;
		for(MGLSummaryObj obj : mglSumList) {
			Map<String, Double[]> map = obj.getMTD();
			for(String key: map.keySet()) {
				int m = 0;
				try {
					m = DateUtil.getMonthNo(key);
				} catch(Exception e) {
					System.err.println("cannot convert: " + key);
					e.printStackTrace();
				}
				max = max < m ? max = new Integer(m) : max;
			}
		}
		return max > 0 ? max + 1 : 1;
	}
	
	private void doTableHeader(Sheet tempSheet, Sheet toSheet, int noOfMonth) throws Exception {
		
		int startRow = new Integer(START_TABLE_HEADER_ROW).intValue();
		int mtdIdx = noOfMonth * 2;
		
		for(int rn = startRow; rn < startRow + 2; rn++) {
			Row tempRow = tempSheet.getRow(rn);
			Row toRow = toSheet.createRow(rn);
			
			Cell tempCell = null;
			Cell toCell = null;
			
			int currMonth = 0;
			int maxCol = tempRow.getLastCellNum() + mtdIdx;
			for(int cn = 0; cn < maxCol; cn++) {
				
//				MTD
				if(cn > 1 && cn < (mtdIdx + 2)) {
					int temp = cn % 2 == 0 ? 2 : 3;
					tempCell = tempRow.getCell(temp, Row.CREATE_NULL_AS_BLANK);
					toCell = toRow.createCell(cn, tempCell.getCellType());
					toCell.setCellStyle(tempCell.getCellStyle());
					WorkbookUtil.getInstance().copyCellValue(tempCell, toCell);
					
					String mmm = toCell.getStringCellValue();
					if(mmm.indexOf(MONTH_PATTERN) > 0) {
						toCell.setCellValue(mmm.replace(MONTH_PATTERN, DateUtil.getStringOfMonth(currMonth)));
						currMonth++;
					}
					WorkbookUtil.getInstance().copyColumnWidth(tempSheet, temp, toSheet, cn);
					
					if(rn == startRow && cn % 2 != 0) {
						CellRangeAddress mergedRegion = new CellRangeAddress(startRow, startRow, cn - 1, cn);
						toSheet.addMergedRegion(mergedRegion);
					}
					
					if(!hideCols.contains(cn)) hideCols.add(new Integer(cn));
				
//				after MTD
				} else if(((cn + 2) - mtdIdx) > 3) {
					int temp = (cn + 2) - mtdIdx;
					
					tempCell = tempRow.getCell(temp, Row.CREATE_NULL_AS_BLANK);
					toCell = toRow.createCell(cn, tempCell.getCellType());
					toCell.setCellStyle(tempCell.getCellStyle());
					WorkbookUtil.getInstance().copyCellValue(tempCell, toCell);
					WorkbookUtil.getInstance().copyColumnWidth(tempSheet, temp, toSheet, cn);
					
					if(rn == startRow && cn % 2 != 0 && temp < 6) {
						CellRangeAddress mergedRegion = new CellRangeAddress(startRow, startRow, cn - 1, cn);
						toSheet.addMergedRegion(mergedRegion);
					} else if(rn == (startRow + 1) && temp > 6) {
						CellRangeAddress mergedRegion = new CellRangeAddress(startRow, rn, cn, cn);
						toSheet.addMergedRegion(mergedRegion);
					}
					
//				campaign
				} else {
					tempCell = tempRow.getCell(cn, Row.CREATE_NULL_AS_BLANK);
					toCell = toRow.createCell(cn, tempCell.getCellType());
					toCell.setCellStyle(tempCell.getCellStyle());
					WorkbookUtil.getInstance().copyCellValue(tempCell, toCell);
					WorkbookUtil.getInstance().copyColumnWidth(tempSheet, cn, toSheet, cn);
					
					if(rn == (startRow + 1)) {
						CellRangeAddress mergedRegion = new CellRangeAddress(startRow, rn, cn, cn);
						toSheet.addMergedRegion(mergedRegion);
					}
				}
			}
		}
	}
	
	private void doTableData(Sheet tempSheet, Sheet toSheet, int maxMonth, List<MGLSummaryObj> mglSumList, Date processDate) throws Exception {
//		remark*: flow is same as table header
		int ytdTarpColIdx = 0, issuedRateColIdx = 0, paidRateColIdx = 0;
		Map<String, Integer[]> colIndexMap = new LinkedHashMap<>();
		int startRow = new Integer(START_TABLE_DATA_ROW).intValue();
		int n = 0;
		int mtdIdx = maxMonth * 2;
		for(int rn = startRow; rn < startRow + mglSumList.size(); rn++) {
			MGLSummaryObj mgl = mglSumList.get(n);
			
//			Double[] sumOfYtd = new Double[]{0D, 0D};
			Row tempRow = tempSheet.getRow(startRow);
			Row toRow = toSheet.createRow(rn);
			
			Cell tempCell = null;
			Cell toCell = null;
			
			int maxCol = tempRow.getLastCellNum() + mtdIdx;
			
			for(int cn = 0; cn < maxCol; cn++) {
				
//				MTD
				if(cn > 1 && cn < (mtdIdx + 2)) {
					boolean isFirstColOfMTD = cn % 2 == 0;
					int temp = isFirstColOfMTD ? 2 : 3;
					double val = 0D;
					
					tempCell = tempRow.getCell(temp, Row.CREATE_NULL_AS_BLANK);
					toCell = toRow.createCell(cn, tempCell.getCellType());
					
					toCell.setCellStyle(tempCell.getCellStyle());
					
					String mtdColMonth = null;
					String monthFromCell = null;
					if(isFirstColOfMTD) {
						monthFromCell = toSheet.getRow(START_TABLE_HEADER_ROW).getCell(cn, Row.CREATE_NULL_AS_BLANK).getStringCellValue();
					} else {
						monthFromCell = toSheet.getRow(START_TABLE_HEADER_ROW).getCell(cn - 1, Row.CREATE_NULL_AS_BLANK).getStringCellValue();
					}
					mtdColMonth = monthFromCell.substring(monthFromCell.indexOf("(") + 1, monthFromCell.indexOf(")"));
					
					Double[] mtdVal = mgl.getMTD().get(mtdColMonth);
					int idx = isFirstColOfMTD ? 0 : 1;
					if(mtdVal != null && mtdVal.length > 0) {
						val = mtdVal[idx].doubleValue();
//						sumOfYtd[idx] += val;
						
//						Double[] mtdByMMM = sumOfMtdMap.get(mtdColMonth);
//						if(mtdByMMM == null) {
//							mtdByMMM = new Double[]{0D, 0D};
//						}
//						mtdByMMM[idx] += val;
//						sumOfMtdMap.put(mtdColMonth, mtdByMMM);
					}
					
					toCell.setCellValue(val);
					
					if(colIndexMap.get(mtdColMonth) == null) {
						colIndexMap.put(mtdColMonth, new Integer[2]);
					}
					Integer[] dataArray = colIndexMap.get(mtdColMonth);
					dataArray[idx] = toCell.getColumnIndex();
//				after MTD
				} else if(((cn + 2) - mtdIdx) > 3) {
					int temp = (cn + 2) - mtdIdx;
					
					tempCell = tempRow.getCell(temp, Row.CREATE_NULL_AS_BLANK);
					toCell = toRow.createCell(cn, tempCell.getCellType());
					toCell.setCellStyle(tempCell.getCellStyle());
					
//					YTD
					if(temp < 6) {
						String columnsFormular = "";
						String sumFormular = "SUM(#COLUMNS)";
						int idx = cn % 2 == 0 ? 0 : 1;
						
						for(String key : colIndexMap.keySet()) {
							Integer colIdex = colIndexMap.get(key)[idx];
							columnsFormular += CellReference.convertNumToColString(colIdex) + (toRow.getRowNum() + 1) + ",";
						}
						toCell.setCellFormula(sumFormular.replace("#COLUMNS", columnsFormular.substring(0, columnsFormular.lastIndexOf(","))));
//						toCell.setCellValue(sumOfYtd[idx]);
//						sumAllOfYTD[idx] += sumOfYtd[idx];
						
						if(cn % 2 != 0) ytdTarpColIdx = toCell.getColumnIndex();
						
					} /* target */ else if(temp > 6) {
						switch(temp) {
						case 7 : 
//							issued rate
							issuedRateColIdx = toCell.getColumnIndex();
							toCell.setCellValue(0); break;
						case 8 : 
//							paid rate
							paidRateColIdx = toCell.getColumnIndex();
							toCell.setCellValue(0); break;
						case 9 : 
//							Double[] mtdVal = mgl.getMTD().get(DateUtil.convDateToString("MMM", processDate));
//							Double iapMTD = mtdVal[1] * mgl.getIssuedRate() * mgl.getPaidRate();
//							toCell.setCellValue(iapMTD);
							
//							<-- set formula -->
							if(hideCols.size() > 0) {
								
								String iapMtdFormula = 
										CellReference.convertNumToColString(hideCols.get(hideCols.size() - 1)) + (rn + 1)
										+ "*"
										+ CellReference.convertNumToColString(issuedRateColIdx) + (rn + 1)
										+ "*"
										+ CellReference.convertNumToColString(paidRateColIdx) + (rn + 1)
										;
								toCell.setCellFormula(iapMtdFormula);
							}
							
//							Double sumIapMtd = sumOfIAP.get(MTD_STR);
//							if(sumIapMtd == null) {
//								sumIapMtd = new Double(0);
//							}
//							sumOfIAP.put(MTD_STR, sumIapMtd + iapMTD);
							break;
						case 10 : 
//							Double iapYTD = toRow.getCell(cn - 5, Row.CREATE_NULL_AS_BLANK).getNumericCellValue() * mgl.getIssuedRate() * mgl.getPaidRate();
//							toCell.setCellValue(iapYTD);
//							Double sumIapYtd = sumOfIAP.get(YTD_STR);
							
							String iapYtdFormula = 
									CellReference.convertNumToColString(ytdTarpColIdx) + (rn + 1)
									+ "*"
									+ CellReference.convertNumToColString(issuedRateColIdx) + (rn + 1)
									+ "*"
									+ CellReference.convertNumToColString(paidRateColIdx) + (rn + 1)
									;
							
							toCell.setCellFormula(iapYtdFormula);
							
//							if(sumIapYtd == null) {
//								sumIapYtd = new Double(0);
//							}
//							sumOfIAP.put(YTD_STR, sumIapYtd + iapYTD);
							break;
						default : break;
						}
					}
					
//				campaign
				} else {
					tempCell = tempRow.getCell(cn, Row.CREATE_NULL_AS_BLANK);
					toCell = toRow.createCell(cn, tempCell.getCellType());
					toCell.setCellStyle(tempCell.getCellStyle());
					
					toCell.setCellValue(cn % 2 == 0 ? mgl.getCampaignName() : mgl.getCampaignCode());
				}
			}
			n++;
		}
	}
	
	private void doTableTotal(Sheet tempSheet, Sheet toSheet, int maxMonth) {
		int startRow = toSheet.getLastRowNum() + 1;
		int mtdIdx = maxMonth * 2;
		String sumFunction = "SUM(#FROM:#TO)";
		Row tempRow = tempSheet.getRow(TEMP_TABLE_TOTAL_ROW);
		Row toRow = toSheet.createRow(startRow);
		
		Cell tempCell = null;
		Cell toCell = null;
		int maxCol = mtdIdx + tempRow.getLastCellNum();
		
		for(int cn = 0; cn < maxCol; cn++) {
			
//			MTD
			if(cn > 1 && cn < (mtdIdx + 2)) {
				boolean isFirstColOfMTD = cn % 2 == 0;
				int temp = isFirstColOfMTD ? 2 : 3;
				
				tempCell = tempRow.getCell(temp, Row.CREATE_NULL_AS_BLANK);
				toCell = toRow.createCell(cn, tempCell.getCellType());
				
				toCell.setCellStyle(tempCell.getCellStyle());
				
//				String mtdColMonth = null;
//				String monthFromCell = null;
//				if(isFirstColOfMTD) {
//					monthFromCell = toSheet.getRow(START_TABLE_HEADER_ROW).getCell(cn, Row.CREATE_NULL_AS_BLANK).getStringCellValue();
//				} else {
//					monthFromCell = toSheet.getRow(START_TABLE_HEADER_ROW).getCell(cn - 1, Row.CREATE_NULL_AS_BLANK).getStringCellValue();
//				}
//				mtdColMonth = monthFromCell.substring(monthFromCell.indexOf("(") + 1, monthFromCell.indexOf(")"));
				
				toCell.setCellFormula(
						sumFunction
							.replace("#FROM", CellReference.convertNumToColString(toCell.getColumnIndex()) + (START_TABLE_DATA_ROW + 1))
							.replace("#TO", CellReference.convertNumToColString(toCell.getColumnIndex()) + toCell.getRowIndex()));
//				Double[] mtd = sumOfMtdMap.get(mtdColMonth);
//				if(mtd != null) {
//					toCell.setCellValue(sumOfMtdMap.get(mtdColMonth)[isFirstColOfMTD ? 0 : 1]);
//				} else {
//					toCell.setCellValue(0);
//				}
				
//			after MTD
			} else if(((cn + 2) - mtdIdx) > 3) {
				int temp = (cn + 2) - mtdIdx;
				
				tempCell = tempRow.getCell(temp, Row.CREATE_NULL_AS_BLANK);
				toCell = toRow.createCell(cn, tempCell.getCellType());
				toCell.setCellStyle(tempCell.getCellStyle());
				
//				YTD
				if(temp < 6) {
//					toCell.setCellValue(sumAllOfYTD[cn % 2 == 0 ? 0 : 1]);
					toCell.setCellFormula(
							sumFunction
								.replace("#FROM", CellReference.convertNumToColString(toCell.getColumnIndex()) + (START_TABLE_DATA_ROW + 1))
								.replace("#TO", CellReference.convertNumToColString(toCell.getColumnIndex()) + toCell.getRowIndex()));
				} /* target */ 
				else if(temp > 6) {
					if(temp == 9 || temp == 10) {
						toCell.setCellFormula(
								sumFunction
									.replace("#FROM", CellReference.convertNumToColString(toCell.getColumnIndex()) + (START_TABLE_DATA_ROW + 1))
									.replace("#TO", CellReference.convertNumToColString(toCell.getColumnIndex()) + toCell.getRowIndex()));
					}
				}
				
//			campaign
			} else {
				tempCell = tempRow.getCell(cn, Row.CREATE_NULL_AS_BLANK);
				toCell = toRow.createCell(cn, tempCell.getCellType());
				toCell.setCellStyle(tempCell.getCellStyle());
				
				if(cn == 0) {
					WorkbookUtil.getInstance().copyCellValue(tempCell, toCell);
				} else {
					toSheet.addMergedRegion(new CellRangeAddress(startRow, startRow, cn - 1, cn));
				}
			}
		}
		
	}
	
	private void writeOut(Workbook wb, Date processDate, String outPath) throws IOException {
		String yyyy = "yyyy";
		String outName = EXPORT_FILE_NAME.replaceAll("#".concat(yyyy), DateUtil.convDateToString(yyyy, processDate));
		FileUtil.getInstance().createDirectory(outPath);
		String outDir = WorkbookUtil.getInstance().writeOut(wb, outPath, outName);
		wb.close();
		wb = null;
		logger.info("Writed to " + outDir);
	}

}
