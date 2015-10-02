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
			String yyyyMMarg = args[0];
			Date processDate = DateUtil.toEndOfMonth(DateUtil.convStringToDate("yyyyMMdd", yyyyMMarg+"01"));
			String dir = args[1];
			
			new MGLSummaryReport().generateReport(dir + File.separatorChar + yyyyMMarg + File.separatorChar + "summary", processDate);
			
			new MGLByCampaignMonthlyReport().generateReport(dir + File.separatorChar + yyyyMMarg + File.separatorChar + "production", processDate);
			
			new PlanLVReport().generateReport(dir + File.separatorChar + yyyyMMarg + File.separatorChar + "planlv", processDate);
			
			logger.info("### Finish ###");
		} catch(Exception e) {
			e.printStackTrace();    
		}
	}
}
