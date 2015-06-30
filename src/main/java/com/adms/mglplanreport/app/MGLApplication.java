package com.adms.mglplanreport.app;

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
			
			String processDateStr = args[0];
//			String dir = "D:/project/reports/MGL/out";
			String dir = args[1];
			Date processDate = DateUtil.convStringToDate("yyyyMMdd", processDateStr);
			
			new MGLSummaryReport().generateReport(dir + "/" + processDateStr.substring(0, 6) + "/summary", processDate);
			
			new MGLByCampaignMonthlyReport().generateReport(dir + "/" + processDateStr.substring(0, 6) + "/production", processDate);
			
			new PlanLVReport().generateReport(dir + "/" + processDateStr.substring(0, 6) + "/planlv", processDate);
			
			logger.info("### Finish ###");
		} catch(Exception e) {
			e.printStackTrace();    
		}
	}
}
