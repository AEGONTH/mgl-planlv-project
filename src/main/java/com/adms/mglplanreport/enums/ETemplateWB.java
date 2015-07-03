package com.adms.mglplanreport.enums;

public enum ETemplateWB {

	MGL_SUMMARY_TEMPLATE("MGL Summary", "template/MGL-Template.xlsx", 0),
	MGL_BY_CAMPAIGN_TEMPLATE("MGL By Campaign", "template/MGL-Template.xlsx", 1),
	MGL_BY_CAMPAIGN_TOTAL_TEMPLATE("MGL By Campaign Total", "template/MGL-Template.xlsx", 2),
	PLAN_LV_TEMPLATE("Production Plan Level", "template/Production-PlanLv-Template.xlsx", 99);
	
	private String name;
	private String filePath;
	private int sheetIndex;
	
	private ETemplateWB(String name, String filePath, int sheetIndex) {
		this.name = name;
		this.filePath = filePath;
		this.sheetIndex = sheetIndex;
	}

	public String getName() {
		return name;
	}

	public String getFilePath() {
		return filePath;
	}

	public int getSheetIndex() {
		return sheetIndex;
	}
	
}
