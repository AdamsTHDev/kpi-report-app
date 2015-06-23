package com.adms.batch.kpireport.app;

import com.adms.batch.kpireport.job.ExportJob;
import com.adms.utils.Logger;

public class ExportKpiReports {
	private static Logger logger = Logger.getLogger();

	public static void main(String[] args) {
		try {
			String processDate = args[0];
			String outPath = args[1];
			
			logger.setLogFileName(args[2]);
			
//			String processDate = "20150531";
//			String outPath = "d:/temp/kpi";
			
			ExportJob.getInstance(processDate).exportKpiReports(outPath);
		} catch(Exception e) {
			logger.error(e.getMessage(), e);
		}
	}
}
