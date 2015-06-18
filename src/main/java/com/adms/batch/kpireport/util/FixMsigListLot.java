package com.adms.batch.kpireport.util;

import java.util.List;

import org.apache.commons.lang3.StringUtils;

import com.adms.entity.KpiCategorySetup;
import com.adms.kpireport.service.KpiCategorySetupService;

public class FixMsigListLot {

	private static FixMsigListLot instance;
	
	private String resultMSIG = "";

	private KpiCategorySetupService kpiCategorySetupService = (KpiCategorySetupService) AppConfig.getInstance().getBean("kpiCategorySetupService");
	
	public static FixMsigListLot getInstance() {
		if(instance == null) {
			instance = new FixMsigListLot();
		}
		return instance;
	}
	
	/**
	 * Getting List lot code of MISG POM WB with delimited ',' ex: AGA15,AGB15
	 * @param processDate as String pattern yyyyMMdd or at least yyyy
	 * @return String with ',' delimited
	 * @throws Exception
	 */
	public String getFixMisgListLotDelim(String processDate) throws Exception {
		if(resultMSIG == null || StringUtils.isBlank(resultMSIG)) {
			
			String hql = " from KpiCategorySetup d "
					+ " where 1 = 1 "
					+ " and d.listLotCode is not null "
					+ " and CONVERT(nvarchar(4), d.effectiveDate, 112) = ? "
					+ " order by d.effectiveDate ";
			List<KpiCategorySetup> list = kpiCategorySetupService.findByHql(hql, processDate.substring(0, 4));
			
			for(KpiCategorySetup k : list) {
				if(!resultMSIG.contains(k.getListLotCode())) {
					if(StringUtils.isBlank(resultMSIG)) {
						resultMSIG += k.getListLotCode();
					} else {
						resultMSIG += "," + k.getListLotCode();
					}
				}
			}
		}
		
		return resultMSIG;
	}
}
