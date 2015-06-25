package com.adms.batch.kpireport.job;

import com.adms.batch.kpireport.service.ReportExporter;
import com.adms.batch.kpireport.service.impl.KpiReportExporter;
import com.adms.utils.Logger;

public class ExportJob {

	private static ExportJob instance;
		
	private String processDate = "";
	
	private static Logger logger = Logger.getLogger();
	
	public static ExportJob getInstance(String processDate) {
		if(instance == null) {
			instance = new ExportJob(processDate);
		}
		return instance;
	}
	
	private ExportJob(String processDate) {
		this.processDate = processDate;
	}
	
	public void exportKpiReports(String destination) {
		logger.info("### Start Export KPI Reports ###");
		
		try {
			ReportExporter export = new KpiReportExporter();
			export.exportExcel(destination, processDate);
		} catch(Exception e) {
			logger.error(e.getMessage(), e);
		}
		
		logger.info("### Finish Export KPI Reports ###");
	}
	
}
