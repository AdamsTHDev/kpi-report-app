package com.adms.batch.kpireport.enums;

public enum EFileFormat {
	TSR_TRACKING_TELE("TSR_TRACKING_TELE", "config/fileformat/TSR_TRACKING_FORMAT_TELE.xml"),
	TSR_TRACKING_OTO("TSR_TRACKING_OTO", "config/fileformat/TSR_TRACKING_FORMAT_OTO.xml"),
	KPI_TARGET_FORMAT("KPI_TARGET_FORMAT", "config/fileformat/KPI_TARGET_FORMAT.xml"),
	SUP_DSM("SUP_DSM_FORMAT", "config/fileformat/SUP_DSM_FORMAT.xml");
	
	private String code;
	private String value;
	
	private EFileFormat(String code, String value) {
		this.code = code;
		this.value = value;
	}

	public String getCode() {
		return code;
	}
	
	public String getValue() {
		return value;
	}
}
