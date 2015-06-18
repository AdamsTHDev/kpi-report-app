package com.adms.batch.kpireport.job;
import java.io.File;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.List;

import org.apache.commons.lang3.StringUtils;

import com.adms.batch.kpireport.service.DataImporter;
import com.adms.batch.kpireport.service.DataImporterFactory;
import com.adms.batch.kpireport.service.impl.KpiTargetSetupImporter;
import com.adms.batch.kpireport.service.impl.SupDsmImporter;
import com.adms.utils.Logger;

public class ImportJob {

	public static ImportJob instance;
	
	private static Logger logger = Logger.getLogger();
	
	private String processDate = "";
	private List<String> dirs;
	
	/**
	 * 
	 * @param processDate must be yyyyMMdd format
	 * @return instance
	 */
	public static ImportJob getInstance(final String processDate) {
		if(instance == null) {
			instance = new ImportJob(processDate);
		}
		return instance;
	}
	
	private ImportJob(String processDate) {
		this.processDate = processDate;
	}
	
	public void importDataForKPI(String dir) {
		logger.info("#### Start Import Data for KPI");
		
		if(StringUtils.isBlank(dir)) return;
		
		logicImportKpi(dir, new String[]{"TsrTracking", "TSRTracking", "TSRTRA"}, new String[]{"CTD", "MTD", "DAI_ALL", "QA_Report", "QC_Reconfirm", "SalesReportByRecords"});
		
		logger.info("#### Finish Import Data for KPI");
	}
	
	public void importSupDsm(String dir) {
		logger.info("#### Start Import SUP DSM");
		
		try {
			DataImporter importer = new SupDsmImporter();
			importer.importData(dir, processDate);
		} catch(Exception e) {
			logger.error(e.getMessage(), e);
		}
		
		logger.info("#### Finish Import SUP DSM");
	}
	
	public void importKpiTargetSetup(String dir) {
		logger.info("#### Start Import KPI Target Setup");
				
		try {
			DataImporter importer = new KpiTargetSetupImporter();
			importer.importData(dir, processDate);
		} catch(Exception e) {
			logger.error(e.getMessage(), e);
		}
		
		logger.info("#### Finish Import KPI Target Setup");
	}
	
	private void logicImportKpi(String root, String[] require, String[] notIn) {
		logger.info("## Filter > " + Arrays.toString(require) + " | and not > " + Arrays.toString(notIn));
		dirs = new ArrayList<>();
		getExcelByName(root, require, notIn);
		
		for(String dir : dirs) {
			logger.info("# do: " + dir);
			DataImporter importer = null;
			try {
				importer = DataImporterFactory.getDataImporter(dir);
				importer.importData(dir);
			} catch(Exception e) {
				logger.error(e.getMessage(), e);
			} finally {
				
			}
		}
		
	}
	
	private void getExcelByName(String rootPath, String[] containNames, String[] notInNames) {
		File file = new File(rootPath);
		if(file.isDirectory()) {
			for(File sub : file.listFiles()) {
				getExcelByName(sub.getAbsolutePath(), containNames, notInNames);
			}
		} else {
			for(String name : containNames) {
				if(file.getName().contains(name)) {
					if(file.getName().toLowerCase().endsWith(".xls") || file.getName().toLowerCase().endsWith(".xlsx")) {
						boolean flag = true;
						if(notInNames != null && notInNames.length > 0) {
							for(String not : notInNames) {
								if(file.getName().contains(not)) {
									flag = false;
									break;
								}
							}
						}
						if(flag) addToDirs(file.getAbsolutePath());
						
					}
				}
			}
		}
	}
	
	private void addToDirs(String dir) {
		if(dirs == null) {
			dirs = new ArrayList<String>();
		}
		dirs.add(dir);
	}
}
