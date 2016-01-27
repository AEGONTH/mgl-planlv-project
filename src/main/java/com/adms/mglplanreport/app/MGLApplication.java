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
			String method = args[3];
			
			if(method.equals("SUMMARY_REPORT")) {
				new MGLSummaryReport().generateReport(dir + File.separatorChar + yyyyMMarg + File.separatorChar + "summary", processDate);
			} else if(method.equals("PRODUCTION_REPORT")) {
				new MGLByCampaignMonthlyReport().generateReport(dir + File.separatorChar + yyyyMMarg + File.separatorChar + "production", processDate);
			} else if(method.equals("PLAN_LV_REPORT")) {
				new PlanLVReport().generateReport(dir + File.separatorChar + yyyyMMarg + File.separatorChar + "planlv", processDate);
			} else {
				logger.error("Method Not Found...");
			}
			
			logger.info("### Finish ###");
		} catch(Exception e) {
			e.printStackTrace();    
		}
	}
}
