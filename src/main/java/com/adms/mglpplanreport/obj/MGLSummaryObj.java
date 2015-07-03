package com.adms.mglpplanreport.obj;

import java.util.Map;

public class MGLSummaryObj {

	private String campaignName;
	private String campaignCode;
	
	private Map<String, Double[]> MTD;
	private Double issuedRate;
	private Double paidRate;
	private Double iapMTD;
	private Double iapYTD;
	
	public String getCampaignName() {
		return campaignName;
	}
	public void setCampaignName(String campaignName) {
		this.campaignName = campaignName;
	}
	public String getCampaignCode() {
		return campaignCode;
	}
	public void setCampaignCode(String campaignCode) {
		this.campaignCode = campaignCode;
	}
	public Map<String, Double[]> getMTD() {
		return MTD;
	}
	public void setMTD(Map<String, Double[]> mTD) {
		MTD = mTD;
	}
	public Double getIssuedRate() {
		return issuedRate;
	}
	public void setIssuedRate(Double issuedRate) {
		this.issuedRate = issuedRate;
	}
	public Double getPaidRate() {
		return paidRate;
	}
	public void setPaidRate(Double paidRate) {
		this.paidRate = paidRate;
	}
	public Double getIapMTD() {
		return iapMTD;
	}
	public void setIapMTD(Double iapMTD) {
		this.iapMTD = iapMTD;
	}
	public Double getIapYTD() {
		return iapYTD;
	}
	public void setIapYTD(Double iapYTD) {
		this.iapYTD = iapYTD;
	}
	
}
