package com.adms.mglpplanreport.obj;

import java.util.List;

import com.adms.mglplanlv.entity.ProductionByLot;

public class MGLMonthlyObj {

	private String month;
	
	private String campaignCode;
	
	private String campaignName;
	
	private String listLot;
	
	private Double rercordsReceived;
	
	private String callingDate;
	
	private String callingSite;
	
	private List<ProductionByLot> productionByLots;

	public String getMonth() {
		return month;
	}

	public void setMonth(String month) {
		this.month = month;
	}

	public String getCampaignCode() {
		return campaignCode;
	}

	public void setCampaignCode(String campaignCode) {
		this.campaignCode = campaignCode;
	}

	public String getCampaignName() {
		return campaignName;
	}

	public void setCampaignName(String campaignName) {
		this.campaignName = campaignName;
	}

	public String getListLot() {
		return listLot;
	}

	public void setListLot(String listLot) {
		this.listLot = listLot;
	}

	public Double getRercordsReceived() {
		return rercordsReceived;
	}

	public void setRercordsReceived(Double rercordsReceived) {
		this.rercordsReceived = rercordsReceived;
	}

	public String getCallingDate() {
		return callingDate;
	}

	public void setCallingDate(String callingDate) {
		this.callingDate = callingDate;
	}

	public String getCallingSite() {
		return callingSite;
	}

	public void setCallingSite(String callingSite) {
		this.callingSite = callingSite;
	}

	public List<ProductionByLot> getProductionByLots() {
		return productionByLots;
	}

	public void setProductionByLots(List<ProductionByLot> productionByLots) {
		this.productionByLots = productionByLots;
	}

	@Override
	public String toString() {
		return "MGLMonthlyObj [month=" + month + ", campaignCode="
				+ campaignCode + ", campaignName=" + campaignName
				+ ", listLot=" + listLot + ", rercordsReceived="
				+ rercordsReceived + ", callingDate=" + callingDate
				+ ", callingSite=" + callingSite + ", productionByLots="
				+ productionByLots + "]";
	}

}
