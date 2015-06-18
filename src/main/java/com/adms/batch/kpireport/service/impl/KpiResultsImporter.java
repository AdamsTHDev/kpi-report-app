package com.adms.batch.kpireport.service.impl;

import java.math.BigDecimal;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

import com.adms.batch.kpireport.service.DataImporter;
import com.adms.batch.kpireport.util.AppConfig;
import com.adms.batch.kpireport.util.FixMsigListLot;
import com.adms.entity.Campaign;
import com.adms.entity.KpiResult;
import com.adms.entity.Tsr;
import com.adms.entity.bean.KpiBean;
import com.adms.kpireport.service.CampaignService;
import com.adms.kpireport.service.KpiBeanService;
import com.adms.kpireport.service.KpiResultService;
import com.adms.kpireport.service.TsrService;
import com.adms.utils.Logger;

public class KpiResultsImporter implements DataImporter {
	
	private Logger logger = Logger.getLogger();
	
	private final String USER_LOGIN = "KPI Importer";
	
	private KpiBeanService kpiBeanService = (KpiBeanService) AppConfig.getInstance().getBean("kpiBeanService");
	
	private KpiResultService kpiResultService = (KpiResultService) AppConfig.getInstance().getBean("kpiResultService");
	
	private CampaignService campaignService = (CampaignService) AppConfig.getInstance().getBean("campaignService");
	
	private TsrService tsrService = (TsrService) AppConfig.getInstance().getBean("tsrService");
	
	@Deprecated
	@Override
	public void importData(String path) throws Exception {
		throw new Exception("THIS Method is not avilable");
	}

	@Override
	public void importData(String path, String processDate) throws Exception {
		addDataToKpiResult(processDate.substring(0, 6));
	}

	private void addDataToKpiResult(String yearMonth) throws Exception {
	
	//		<!-- Delete old data -->
			logger.info("## Clearing Old Data by YearMonth: " + yearMonth + " ##");
			String hql = "delete from KpiResult where yearMonth = ?";
			int n = kpiResultService.deleteByHql(hql, yearMonth);
			logger.info("# Total: " + n + " records");
		
			logger.info("## Adding data to KPI Result ##");
	//		<!-- get data -->
			List<KpiBean> kpiBeanList = kpiBeanService.findByNamedQuery("getKpiBeanByDate", yearMonth);
			
	//		<!-- Add to Kpi Result -->
			logicKpiResult(kpiBeanList, yearMonth, FixMsigListLot.getInstance().getFixMisgListLotDelim(yearMonth));
		}

	private void logicKpiResult(List<KpiBean> kpiBeans, String yearMonth, String delimKeyCodeMSIGWB) throws Exception {
		try {
			
			Map<String, Map<String, Map<String, Map<String, KpiResult>>>> campaignMap = new HashMap<>();
			Map<String, Map<String, Map<String, KpiResult>>> dsmMap = null;
			Map<String, Map<String, KpiResult>> tsmMap = null;
			Map<String, KpiResult> tsrMap = null;
	
			String campaignCode = "";
			String listLotCode = "";
			
			String dsmCode = "";
			String tsmCode = "";
			String tsrCode = "";
			
			KpiResult kpiResult = null;
			
			BigDecimal afyp = null;
			Integer talkDate = null;
			BigDecimal talkHrs = null;
			Integer fcs = null;
			Integer sale = null;
			Integer successEoc = null;
			Integer allEoc = null;
			
			for(KpiBean kpiBean : kpiBeans) {
				
				if(delimKeyCodeMSIGWB.contains(kpiBean.getListLotCode())) {
					listLotCode = delimKeyCodeMSIGWB.split(",")[delimKeyCodeMSIGWB.split(",").length - 1];
					campaignCode = new String(kpiBean.getCampaignCode() + "_" + listLotCode);
				} else {
					campaignCode = new String(kpiBean.getCampaignCode());
					listLotCode = null;
				}
				
				dsmCode = kpiBean.getDsmCode();
				tsmCode = kpiBean.getSupCode();
				tsrCode = kpiBean.getTsrCode();
	
				dsmMap = campaignMap.get(campaignCode);
				if(dsmMap == null) {
					dsmMap = new HashMap<>();
					campaignMap.put(campaignCode, dsmMap);
				}
				
				tsmMap = dsmMap.get(dsmCode);
				if(tsmMap == null) {
					tsmMap = new HashMap<>();
					dsmMap.put(dsmCode, tsmMap);
				}
				
				tsrMap = tsmMap.get(tsmCode);
				if(tsrMap == null) {
					tsrMap = new HashMap<>();
					tsmMap.put(tsmCode, tsrMap);
				}
				
				kpiResult = tsrMap.get(tsrCode);
				if(kpiResult == null) {
					kpiResult = new KpiResult();
					
					kpiResult.setYearMonth(yearMonth);
					kpiResult.setCampaign(campaignService.find(new Campaign(kpiBean.getCampaignCode())).get(0));
					kpiResult.setDsm(tsrService.find(new Tsr(dsmCode)).get(0));
					kpiResult.setTsm(tsrService.find(new Tsr(tsmCode)).get(0));
					kpiResult.setTsr(tsrService.find(new Tsr(tsrCode)).get(0));
					
					kpiResult.setListLotCode(listLotCode);
					
					afyp = (kpiBean.getTotalAfyp() == null ? new BigDecimal(0) : kpiBean.getTotalAfyp());
					talkDate = (kpiBean.getAttendance() == null ? new Integer(0) : kpiBean.getAttendance());
					talkHrs = (kpiBean.getTotalTalkTime() == null ? new BigDecimal(0) : kpiBean.getTotalTalkTime());
					fcs = (kpiBean.getFirstConfirm() == null ? new Integer(0) : kpiBean.getFirstConfirm().intValue());
					sale = (kpiBean.getTotalApp() == null ? new Integer(0) : kpiBean.getTotalApp().intValue());
					successEoc = (kpiBean.getSuccessPolicy() == null ? new Integer(0) : kpiBean.getSuccessPolicy());
					allEoc = (kpiBean.getTotalUsed() == null ? new Integer(0) : kpiBean.getTotalUsed());
					
				} else {
					afyp = (kpiBean.getTotalAfyp() == null ? new BigDecimal(0) : kpiBean.getTotalAfyp()).add(kpiResult.getTotalAfyp());
					talkDate = (kpiBean.getAttendance() == null ? new Integer(0) : kpiBean.getAttendance());
					talkHrs = (kpiBean.getTotalTalkTime() == null ? new BigDecimal(0) : kpiBean.getTotalTalkTime());
					fcs = (kpiBean.getFirstConfirm() == null ? new Integer(0) : kpiBean.getFirstConfirm().intValue()) + kpiResult.getFirstConfirmSale().intValue();
					sale = (kpiBean.getTotalApp() == null ? new Integer(0) : kpiBean.getTotalApp().intValue()) + kpiResult.getAllSale();
					successEoc = (kpiBean.getSuccessPolicy() == null ? new Integer(0) : kpiBean.getSuccessPolicy()) + kpiResult.getSuccessPolicy();
					allEoc = (kpiBean.getTotalUsed() == null ? new Integer(0) : kpiBean.getTotalUsed()) + kpiResult.getTotalUsed();
				}
	
				kpiResult.setTotalAfyp(afyp);
				kpiResult.setAttendance(talkDate);
				kpiResult.setTotalTalkHrs(talkHrs);
				kpiResult.setFirstConfirmSale(fcs);
				kpiResult.setAllSale(sale);
				kpiResult.setSuccessPolicy(successEoc);
				kpiResult.setTotalUsed(allEoc);
				
				tsrMap.put(tsrCode, kpiResult);
				
			}
			
			addKpiResultFromMap(campaignMap);
			
		} catch(Exception e) {
			throw e;
		}
	}

	private void addKpiResultFromMap(Object obj) throws Exception {
		if(obj instanceof Map<?, ?>) {
			for(Object key : ((Map<?, ?>) obj).keySet()) {
				addKpiResultFromMap(((Map<?, ?>) obj).get(key));
			}
		} else if(obj instanceof KpiResult) {
			addKpiResult((KpiResult) obj);
		} else {
			throw new Exception("Not found instance of " + obj.getClass());
		}
	}

	private KpiResult addKpiResult(KpiResult kpiResult) throws Exception {
		try {
			return kpiResultService.add(kpiResult, USER_LOGIN);
		} catch(Exception e) {
			throw e;
		}
	}
}
