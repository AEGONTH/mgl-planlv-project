package com.adms.mglplanreport.service.planlv.factory;

import com.adms.mglplanreport.service.planlv.PlanLevelGenerator;
import com.adms.mglplanreport.service.planlv.impl.AllPlanLvGenImpl;

public class PlanLevelGeneratorFactory {

	public static PlanLevelGenerator getGenerator(String campaign) throws Exception {
		
		return new AllPlanLvGenImpl();
		
//		if(campaign.toUpperCase().contains("MTLIFE KBANK")) {
//			return new MTLKbankPlanLvGenImpl();
//			
//		} else if(campaign.toUpperCase().contains("MTLIFE BROKER")
//				|| campaign.contains("V-Club")) {
//			return new MTLBrokerPlanLvGenImpl();
//			
//		} else if(campaign.toUpperCase().contains("MSIG BROKER")) {
//			return new MSIGBrokerPlanLvGenImpl();
//			
//		} else if(campaign.toUpperCase().contains("MTI KBANK")) {
//			return new MTIKBankPlanLvGenImpl();
//			
//		} else if(campaign.toUpperCase().contains("MSIG UOB")) {
//			return new MSIGUOBPlanLvGenImpl();
//			
//		} else if(campaign.toUpperCase().contains("FWD TVD")) {
//			return new FWDTVDPlanLvGenImpl();
//			
//		} else if(campaign.toUpperCase().contains("GEN HEALTH")) {
//			return new GenHealthPlanLvGenImpl();
//			
//		} else {
//			throw new Exception("Class for: \"" + campaign + "\" not found");
//		}
		
	}
}
