package com.adms.mglplanreport.service.planlv;

import java.util.Date;

import org.apache.poi.ss.usermodel.Sheet;

import com.adms.mglpplanreport.obj.PlanLevelObj;

public interface PlanLevelGenerator {

	public PlanLevelObj getMTDData(String campaignCode, Date processDate) throws Exception;
	public PlanLevelObj getYTDData(String campaignCode, Date processDate) throws Exception;
	public void generateDataSheet(Sheet tempSheet, PlanLevelObj planLevelMtdObj, PlanLevelObj planLevelYTDObj) throws Exception;
}
