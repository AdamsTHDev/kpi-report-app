package com.adms.batch.kpireport.service;

import java.io.File;

import com.adms.batch.kpireport.service.impl.TsrTrackingImporter;

public class DataImporterFactory {

	public static DataImporter getDataImporter(String fullPath) {
		File f = new File(fullPath);
		if((fullPath.contains("TsrTracking")
				|| fullPath.contains("TSRTracking")
				|| fullPath.contains("TSRTRA")
				)
				&& !f.getName().startsWith("QA") 
				&& !f.getName().startsWith("QC")
				&& !fullPath.contains("~")
				) {
			return new TsrTrackingImporter();
		}
		
		return null;
	}
}
