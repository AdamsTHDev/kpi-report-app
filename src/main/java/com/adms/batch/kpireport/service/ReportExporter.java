package com.adms.batch.kpireport.service;

public interface ReportExporter {

	public void exportExcel(String destination, String processDate) throws Exception;
}
