package com.adms.test;

import com.adms.batch.kpireport.job.ExportJob;

public class TestApp {
	
	public static void main(String[] args) {
		
		ExportJob.getInstance("20150531").exportKpiReports("d:/temp/kpi");
	}
}

