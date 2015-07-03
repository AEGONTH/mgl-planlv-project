package com.adms.mglpplanreport.obj;

import java.io.Serializable;
import java.util.List;

import com.adms.mglplanlv.entity.PlanLvValue;

public class PlanLevelObj implements Serializable {
	
	private static final long serialVersionUID = -7387444580666082646L;

	private String campaignCode;
	
	private String monthYear;
	
	private List<PlanLvValue> planLvValues;

	public String getCampaignCode() {
		return campaignCode;
	}

	public void setCampaignCode(String campaignCode) {
		this.campaignCode = campaignCode;
	}

	public String getMonthYear() {
		return monthYear;
	}

	public void setMonthYear(String monthYear) {
		this.monthYear = monthYear;
	}

	public List<PlanLvValue> getPlanLvValues() {
		return planLvValues;
	}

	public void setPlanLvValues(List<PlanLvValue> planLvValues) {
		this.planLvValues = planLvValues;
	}

}
