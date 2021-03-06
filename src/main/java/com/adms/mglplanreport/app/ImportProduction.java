package com.adms.mglplanreport.app;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FilenameFilter;
import java.io.IOException;
import java.io.InputStream;
import java.math.BigDecimal;
import java.net.URLClassLoader;
import java.util.Date;
import java.util.List;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.ss.util.CellReference;

import com.adms.imex.excelformat.DataHolder;
import com.adms.imex.excelformat.ExcelFormat;
import com.adms.mglplanlv.entity.ListLot;
import com.adms.mglplanlv.entity.ProductionByLot;
import com.adms.mglplanlv.service.listlot.ListLotService;
import com.adms.mglplanlv.service.productionbylot.ProductionByLotService;
import com.adms.mglplanreport.util.ApplicationContextUtil;
import com.adms.support.FileWalker;
import com.adms.utils.Logger;

public class ImportProduction {
	
	public static final String BATCH_ID = "BATCH_ID";
//	public static final String APPLICATION_CONTEXT_FILE = "config/application-context.xml";

	public static final String FILE_FORMAT_PRODUCTION_BY_LOT_TELE = "fileformat/FileFormat_SSIS_DailyProductionByLot-input-TELE.xml";
	public static final String FILE_FORMAT_PRODUCTION_BY_LOT_OTO = "fileformat/FileFormat_SSIS_DailyProductionByLot-input-OTO.xml";
	public static final String FILE_FORMAT_PRODUCTION_BY_LOT_3RD = "fileformat/FileFormat_SSIS_DailyProductionByLot-input-3RD.xml";

	public static final String PRODUCTION_BY_LOT_SERVICE_BEAN = "productionByLotService";
	public static final String LIST_LOT_SERVICE_BEAN = "listLotService";
	
	private static boolean uobTimeFormat;
	
//	public static String NEW_HH_MM_CAMPAIGN = "";

//	private ApplicationContext applicationContext;
	protected Logger log = Logger.getLogger(Logger.DEBUG);

	protected void setLogLevel(int logLevel)
	{
		this.log.setLogLevel(logLevel);
	}
	
	protected void setLogFileName(String fileName) {
		try {
			this.log.setLogFileName(fileName);
		} catch (FileNotFoundException e) {
			log.error(e.getMessage(), e);
		}
	}


//	protected Object getBean(String beanId)
//	{
//		if (this.applicationContext == null)
//		{
//			this.applicationContext = new ClassPathXmlApplicationContext(APPLICATION_CONTEXT_FILE);
//		}
//
//		return this.applicationContext.getBean(beanId);
//	}

	protected ProductionByLotService getProductionByLotService()
	{
		return (ProductionByLotService) ApplicationContextUtil.getApplicationContext().getBean(PRODUCTION_BY_LOT_SERVICE_BEAN);
	}

	protected ListLotService getListLotService()
	{
		return (ListLotService) ApplicationContextUtil.getApplicationContext().getBean(LIST_LOT_SERVICE_BEAN);
	}
	
//	private BigDecimal convertTimeToBase100(BigDecimal time) {
//		if(time.doubleValue() > 0d) {
//			double full = time.setScale(2, BigDecimal.ROUND_HALF_UP).doubleValue();
//			double r = (full %1d * 100d) / 60d * 100d;
//			return new BigDecimal(new BigDecimal(full / 1).setScale(0, BigDecimal.ROUND_DOWN).toString() + "." + new BigDecimal(r).setScale(0, BigDecimal.ROUND_HALF_UP).toString());
//		}
//		return new BigDecimal(0);
//	}

	private ProductionByLot extractRecord(DataHolder dataHolder, ProductionByLot productionByLot, boolean isNewHHmm) throws Exception
	{
//		log.debug("extractRecord " + dataHolder.printValues());

//		<!-- New HH.MM -->
		BigDecimal minutes = dataHolder.get("minutes").getDecimalValue().setScale(2, BigDecimal.ROUND_HALF_UP);
		String minutesTxt = minutes.toString();
		
		String hoursTxt = null;
		if(uobTimeFormat) {
			BigDecimal hours = dataHolder.get("hours").getDecimalValue().setScale(2, BigDecimal.ROUND_HALF_UP);
			hoursTxt = hours.toString();
			
			productionByLot.setHour(Long.valueOf(hoursTxt.split("\\.")[0]));
			productionByLot.setMinute(Long.valueOf(minutesTxt.split("\\.")[0]));
			productionByLot.setSecond(Long.valueOf(minutesTxt.split("\\.")[1]));
		} else {
			productionByLot.setHour(Long.valueOf(minutesTxt.split("\\.")[0]) / 60);
			productionByLot.setMinute(Long.valueOf(minutesTxt.split("\\.")[0]) % 60);
			
			if(isNewHHmm) {
				productionByLot.setSecond(Long.valueOf(minutesTxt.split("\\.")[1]));
			} else {
				productionByLot.setSecond(Long.valueOf(Math.round(Float.valueOf("0." + minutesTxt.split("\\.")[1]) * 60)));
			}
		}
		
		productionByLot.setDialing(Long.valueOf(dataHolder.get("dialing").getIntValue()));
		productionByLot.setCompleted(Long.valueOf(dataHolder.get("completed").getIntValue()));
		productionByLot.setContact(Long.valueOf(dataHolder.get("contact").getIntValue()));
		productionByLot.setSales(Long.valueOf(dataHolder.get("sales").getIntValue()));
		productionByLot.setAbandons(Long.valueOf(dataHolder.get("abandons").getIntValue()));
		productionByLot.setUwReleaseSales(Long.valueOf(dataHolder.get("uwReleaseSales").getIntValue()));
		productionByLot.setTyp(dataHolder.get("typ").getDecimalValue().setScale(14, BigDecimal.ROUND_HALF_UP));
		productionByLot.setTotalCost(dataHolder.get("totalCost").getDecimalValue().setScale(14, BigDecimal.ROUND_HALF_UP));
		productionByLot.setReleaseSales(Long.valueOf(dataHolder.get("releaseSales").getIntValue()));
		productionByLot.setAmpPostUw(dataHolder.get("ampPostUw").getDecimalValue().setScale(14, BigDecimal.ROUND_HALF_UP));
		productionByLot.setDeclines(Long.valueOf(dataHolder.get("declines").getIntValue()));

		return productionByLot;
	}
	
	private void importDataHolder(ListLot listLot, Integer totalLead, Integer remainingLead, DataHolder dataHolder, boolean isNewTimeFormat)
			throws Exception
	{
		Date productionDate = (Date) dataHolder.get("productionDate").getValue();
		ProductionByLot productionByLot = getProductionByLotService().findProductionByLotByListLotCodeAndProductionDate(listLot.getListLotCode(), productionDate);

		boolean newProductionByLot = false;
		if (productionByLot == null)
		{
			productionByLot = new ProductionByLot();
			newProductionByLot = true;
		}
		else
		{
//			log.debug("found productionByLot record [" + productionByLot + "]");
		}

		productionByLot.setProductionDate(productionDate);
		productionByLot.setListLot(listLot);
		productionByLot.setTotalLead(Long.valueOf(totalLead));
		productionByLot.setRemainingLead(Long.valueOf(remainingLead));

		try
		{
			extractRecord(dataHolder, productionByLot, isNewTimeFormat);

			if (newProductionByLot)
			{
				getProductionByLotService().addProductionByLot(productionByLot, BATCH_ID);
			}
			else
			{
				getProductionByLotService().updateProductionByLot(productionByLot, BATCH_ID);
			}
		}
		catch (Exception e)
		{
			e.printStackTrace();
		}
	}

	private void importDataHolderList(ListLot listLot, Integer totalLead, Integer remainingLead, List<DataHolder> dataHolderList, boolean isNewTimeFormat)
			throws Exception
	{
		log.info("Importing... ListLot: " + listLot.getListLotCode());
		for (DataHolder dataHolder : dataHolderList)
		{
			importDataHolder(listLot, totalLead, remainingLead, dataHolder, isNewTimeFormat);
		}
	}
	
	private boolean isNewTimeFormat(String filePath) throws InvalidFormatException, FileNotFoundException, IOException {
		boolean flag = false;
		
		Workbook wb = WorkbookFactory.create(new FileInputStream(filePath));
		Row row = wb.getSheetAt(0).getRow(8);
		Cell cell = row.getCell(CellReference.convertColStringToIndex("D"), Row.CREATE_NULL_AS_BLANK);
		String val = cell.getStringCellValue();
		
		if(val.contains("mm.ss")) {
			flag = true;
		}
		wb.close();
		return flag;
	}

	private void importFile(String fileFormatFileName, String dataFileLocation)
	{
		log.info("importFile: " + dataFileLocation);
		InputStream format = null;
		InputStream input = null;
		boolean isNewTimeFormat = false;
		try
		{
			format = URLClassLoader.getSystemResourceAsStream(fileFormatFileName);
			ExcelFormat excelFormat = new ExcelFormat(format);
			
			input = new FileInputStream(dataFileLocation);
			DataHolder fileDataHolder = excelFormat.readExcel(input);

			List<String> sheetNames = fileDataHolder.getKeyList();

			for (String sheetName : sheetNames)
			{
				DataHolder sheetDataHolder = fileDataHolder.get(sheetName);

				ListLot listLot = null;
				Integer totalLead = null;
				Integer remainingLead = null;
				List<DataHolder> listLotList = sheetDataHolder.getDataList("listLotList");
				if (listLotList.size() != 1)
				{
					throw new Exception("listLotCode invalid on sheetName: " + sheetName);
				}
				else
				{
					DataHolder listLotDataHolder = listLotList.get(0);
					String listLotCode = listLotDataHolder.get("listLotCode").getStringValue();
					listLot = getListLotService().findListLotByListLotCode(listLotCode);
					if (listLot == null)
					{
						log.warn("not found listLot for listLotCode: " + listLotCode + " | on sheet: " + sheetName);
						continue;
					}
				}

				List<DataHolder> totalLeadList = sheetDataHolder.getDataList("totalLeadList");
				if (totalLeadList.size() != 1)
				{
					throw new Exception("totalLead invalid on sheetName: " + sheetName);
				}
				else
				{
					DataHolder totalLeadDataHolder = totalLeadList.get(0);
					totalLead = totalLeadDataHolder.get("totalLead").getIntValue();
				}
				
				List<DataHolder> remainingLeadList = sheetDataHolder.getDataList("remainingLeadList");
				if (remainingLeadList.size() != 1)
				{
					throw new Exception("remainingLead invalid on sheetName: " + sheetName);
				}
				else
				{
					DataHolder remainingLeadDataHolder = remainingLeadList.get(0);
					remainingLead = remainingLeadDataHolder.get("remainingLead").getIntValue();
				}

				//TODO Temporary for revise TELE UOB Data
				uobTimeFormat = false;
				
				//TODO Temporary force 3RD using new time format (hh.mm)
				if(dataFileLocation.contains("3RD")) {
					isNewTimeFormat = true;
				} else if(dataFileLocation.contains("TELE") && dataFileLocation.contains("revise")) {
					uobTimeFormat = true;
					isNewTimeFormat = isNewTimeFormat(dataFileLocation);
				} else {
					isNewTimeFormat = isNewTimeFormat(dataFileLocation);
				}
				List<DataHolder> dataHolderList = sheetDataHolder.getDataList("dataRecords");
				importDataHolderList(listLot, totalLead, remainingLead, dataHolderList, isNewTimeFormat);
			}
		}
		catch (Exception e)
		{
			log.error(e.getMessage() + ": " + e.getStackTrace());
		}
		finally
		{
			try
			{
				format.close();
			}
			catch (Exception e)
			{
			}
			try
			{
				input.close();
			}
			catch (Exception e)
			{
			}
		}
	}

	public static void main(String[] args)
	{
		String fileFormatFileName = /* args[0]; */ null;
//		String rootPath = /* args[1]; */ "T:/Business Solution/Share/N_Mos/Daily Sales Report/201502";
//		String rootPath = "T:/Business Solution/Share/AutomateReport/MglReportData/production_report/201506/OTO/finance/MTL/MTL_BL";
		String rootPath = args[0];
		String logPath = args[1];
//		String logPath = "D:/temp/log.log";
//		NEW_HH_MM_CAMPAIGN = args[2];
//		NEW_HH_MM_CAMPAIGN = "021DP1715M01";
		
		ImportProduction batch = new ImportProduction();
		batch.setLogLevel(Logger.DEBUG);
		
		batch.setLogFileName(logPath);
		FileWalker fw = new FileWalker();
		fw.walk(rootPath, new FilenameFilter()
		{
			public boolean accept(File dir, String name)
			{
				return !name.contains("~$") 
						&& !dir.getAbsolutePath().contains("archive")
						&& !name.contains("TSR") 
						&& !name.contains("SalesReportByRecords") 
						&& !name.contains("_ALL")
						&& (name.contains("Production.xls") 
								|| name.contains("Production Report.xlsx") 
								|| name.contains("Production_Report")
								|| (name.contains("Production Report")
										&& name.toLowerCase().contains(".xls"))
								|| name.contains("PRODUC"));
			}
		});

		for (String filename : fw.getFileList())
		{
			if (filename.contains("TELE"))
			{
				fileFormatFileName = FILE_FORMAT_PRODUCTION_BY_LOT_TELE;
			}
			else if (filename.contains("OTO"))
			{
				fileFormatFileName = FILE_FORMAT_PRODUCTION_BY_LOT_OTO;
			}
			else if (filename.contains("3RD")) 
			{
				fileFormatFileName = FILE_FORMAT_PRODUCTION_BY_LOT_3RD;
			} 
			else 
			{
				System.err.println("File format not found for: " + filename);
			}

			batch.importFile(fileFormatFileName, filename);
		}
		System.out.println("Finished");
	}
}
