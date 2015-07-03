package com.adms.mglplanreport.app;

import java.io.File;
import java.util.Date;

import com.adms.mglplanreport.service.MGLByCampaignMonthlyReport;
import com.adms.mglplanreport.service.MGLSummaryReport;
import com.adms.mglplanreport.service.PlanLVReport;
import com.adms.utils.DateUtil;
import com.adms.utils.Logger;

public class MGLApplication {

	private static Logger logger = Logger.getLogger();
	
	public static void main(String[] args) {
		try {
			logger.setLogFileName(args[2]);
//			logger.setLogFileName("d:/temp/log.log");
			
			String processDateStr = args[0];
//			String processDateStr = "20150630";
			Date processDate = DateUtil.convStringToDate("yyyyMMdd", processDateStr);
			
//			String dir = "D:/temp/MGL/out";
			String dir = args[1];
			
			new MGLSummaryReport().generateReport(dir + File.separatorChar + processDateStr.substring(0, 6) + File.separatorChar + "summary", processDate);
			
			new MGLByCampaignMonthlyReport().generateReport(dir + File.separatorChar + processDateStr.substring(0, 6) + File.separatorChar + "production", processDate);
			
			new PlanLVReport().generateReport(dir + File.separatorChar + processDateStr.substring(0, 6) + File.separatorChar + "planlv", processDate);
			
			logger.info("### Finish ###");
		} catch(Exception e) {
			e.printStackTrace();    
		}
	}
}
