package com.adms.batch.kpireport.app;

import com.adms.batch.kpireport.job.ImportJob;
import com.adms.utils.Logger;

public class ImportTracking {
	
	private static Logger logger = Logger.getLogger();

	public static void main(String[] args) {
		try {
			String processDate = args[0];
			String rootPath = args[1];
			logger.setLogFileName(args[2]);
			ImportJob.getInstance(processDate).importDataForKPI(rootPath);
		} catch(Exception e) {
			logger.error(e.getMessage(), e);
		}
	}
}
