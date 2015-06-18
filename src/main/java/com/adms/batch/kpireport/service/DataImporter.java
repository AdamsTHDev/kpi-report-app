package com.adms.batch.kpireport.service;


public interface DataImporter {

	public void importData(final String path) throws Exception;
	public void importData(final String path, final String processDate) throws Exception;
}
