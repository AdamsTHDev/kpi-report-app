package com.adms.batch.kpireport.service.impl;

import java.io.FileInputStream;
import java.io.InputStream;
import java.math.BigDecimal;
import java.util.Arrays;
import java.util.Date;
import java.util.List;

import org.apache.commons.lang3.StringUtils;
import org.hibernate.criterion.DetachedCriteria;
import org.hibernate.criterion.MatchMode;
import org.hibernate.criterion.Restrictions;

import com.adms.batch.kpireport.enums.EFileFormat;
import com.adms.batch.kpireport.service.DataImporter;
import com.adms.batch.kpireport.util.AppConfig;
import com.adms.entity.ListLot;
import com.adms.entity.Tsr;
import com.adms.entity.TsrTracking;
import com.adms.imex.excelformat.DataHolder;
import com.adms.imex.excelformat.ExcelFormat;
import com.adms.kpireport.service.ListLotService;
import com.adms.kpireport.service.TsrService;
import com.adms.kpireport.service.TsrTrackingService;
import com.adms.utils.DateUtil;
import com.adms.utils.Logger;

public class TsrTrackingImporter implements DataImporter {
	
	private static Logger logger = Logger.getLogger();
	
	private final String LOGIN_USER = "TSR_TRACKING_IMPORTER";
	
	private TsrService tsrService = (TsrService) AppConfig.getInstance().getBean("tsrService");
	private TsrTrackingService tsrTrackingService = (TsrTrackingService) AppConfig.getInstance().getBean("tsrTrackingService");
	private ListLotService listLotService = (ListLotService) AppConfig.getInstance().getBean("listLotService");

	private final List<String> titles = Arrays.asList(new String[]{"นาย", "น.ส.", "นาง", "ว่าที่", "นส.", "นางสาว"});
	
	@Override
	public void importData(final String path) throws Exception {
		InputStream fileformatStream = null;
		InputStream wbStream = null;
		ExcelFormat ef = null;
		
		try {
			fileformatStream = Thread.currentThread().getContextClassLoader().getResourceAsStream(this.getFileFormatPath(path));
			wbStream = new FileInputStream(path);
			
			ef = new ExcelFormat(fileformatStream);
			
			DataHolder wbHolder = null;
			wbHolder = ef.readExcel(wbStream);
			
			List<String> sheetNames = wbHolder.getKeyList();
			if(sheetNames.size() == 0) {
				return;
			}
			
			for(String sheetName : sheetNames) {
				logic(wbHolder.get(sheetName), sheetName);
			}
			
		} catch(Exception e) {
			logger.error(e.getMessage(), e);
		} finally {
			try { fileformatStream.close(); } catch(Exception e) {}
			try { wbStream.close(); } catch(Exception e) {}
		}
	}
	
	private void logic(DataHolder sheetHolder, String sheetName) {
		try {
			String period = sheetHolder.get("period").getStringValue().trim().substring(0, 10);
			String listLotName = sheetHolder.get("listLotName").getStringValue();
			List<DataHolder> datas = sheetHolder.getDataList("tsrTrackingList");
			
//			<!-- getting Listlot -->
			if(listLotName.contains(",")) { logger.info("SKIP>> sheetName:" + sheetName + " | listLotName: " + listLotName + " | period: " + period); return;}
			
			String listLotCode = getListLotCode(listLotName);
			if(StringUtils.isBlank(listLotCode)) throw new Exception("Cannot get listlot >> " + listLotName + " | period: " + period);
			
			if(datas.isEmpty()) return;
			
			for(DataHolder data : datas) {
				try {
					
//					<!-- Checking work day -->
					Integer workday = Integer.valueOf(data.get("workday").getDecimalValue() != null 
							? Integer.valueOf(data.get("workday").getDecimalValue().intValue()) : new Integer(0));
					
					if(workday == 0 || workday > 1) {
//						logger.info("## Skip >> workday: " + workday + " | sheetName: " + sheetName);
						continue;
					}
					
//					<!-- getting data -->
					String tsrName = data.get("tsrName").getStringValue();
					if(StringUtils.isEmpty(tsrName)) continue;
					
					Integer listUsed = data.get("listUsed") != null ? data.get("listUsed").getIntValue() : new Integer(0);
					Integer complete = data.get("complete") != null ? data.get("complete").getIntValue() : new Integer(0);
//					listUsed, complete
					
					BigDecimal hours = data.get("hours") != null ? data.get("hours").getDecimalValue().setScale(14, BigDecimal.ROUND_HALF_UP) : new BigDecimal(0);
					BigDecimal talkTime = new BigDecimal(data.get("totalTalkTime").getStringValue()).setScale(14, BigDecimal.ROUND_HALF_UP);
					
					Integer newUsed = data.get("newUsed") != null ? data.get("newUsed").getIntValue() : 0;
					Integer totalPolicy = data.get("totalPolicy") != null ? data.get("totalPolicy").getIntValue() : 0;
					
//					<!-- get Tsr by name -->
					String fullName = removeTitle(tsrName.replaceAll(" ", "").replaceAll("  ", " "));
					Tsr tsr = getTsrByName(fullName);
					if(tsr == null) logger.info("Not found TSR: '" + fullName + "' | listLot: " + listLotCode);
					
					Date trackingDate = null;
					try {
						trackingDate = DateUtil.convStringToDate(period);
					} catch(Exception e) {
						logger.info("Cannot convert String to date > " + period + " | with pattern > " + DateUtil.getDefaultDatePattern());
						logger.info("Try another pattern > dd-MM-yyyy");
						trackingDate = DateUtil.convStringToDate("dd-MM-yyyy", period);
					}
					
					saveTsrTracking(tsr, fullName, trackingDate, listLotCode, workday, listUsed, complete, hours, talkTime, newUsed, totalPolicy);
					
				} catch(Exception e) {
					logger.error(e.getMessage(), e);
				}
			}
			
		} catch(Exception e) {
			logger.error(e.getMessage(), e);
		}
	}
	
	private void saveTsrTracking(Tsr tsr, String tsrName, Date trackingDate, String listLotCode, Integer workday, Integer listUsed, Integer complete, BigDecimal hours, BigDecimal talkTime, Integer newUsed, Integer totalPolicy) throws Exception {
		DetachedCriteria isExisted = DetachedCriteria.forClass(TsrTracking.class);
		isExisted.add(Restrictions.eq("trackingDate", trackingDate));
		isExisted.add(Restrictions.eq("listLot.listLotCode", listLotCode));

		isExisted.add(Restrictions.disjunction()
				.add(Restrictions.eq("tsrName", tsrName))
				.add(Restrictions.eq("tsr.tsrCode", tsr == null ? "" : tsr.getTsrCode())));

		List<TsrTracking> list = tsrTrackingService.findByCriteria(isExisted);
		
		TsrTracking tsrTracking = null;
		if(list.isEmpty()) {
			tsrTracking = new TsrTracking();
			tsrTracking.setTrackingDate(trackingDate);
			tsrTracking.setTsr(tsr);
			tsrTracking.setTsrName(tsrName);
			tsrTracking.setWorkDays(workday);
			tsrTracking.setListUsed(listUsed);
			tsrTracking.setComplete(complete);
			tsrTracking.setListLot(listLotService.find(new ListLot(listLotCode)).get(0));
			tsrTracking.setWorkHours(hours.doubleValue());
			tsrTracking.setTotalTalkTime(talkTime.doubleValue());
			tsrTracking.setNewUsed(newUsed);
			tsrTracking.setTotalPolicy(totalPolicy);
			
			tsrTrackingService.add(tsrTracking, LOGIN_USER);
		} else if(list.size() == 1) {
			boolean isUpdate = false;
			
			tsrTracking = list.get(0);
			if(tsrTracking.getTsr() == null && tsr != null) {
				tsrTracking.setTsr(tsr);
				isUpdate = true;
			}
			if(!tsrTracking.getListUsed().equals(listUsed)) {
				tsrTracking.setListUsed(listUsed);
				isUpdate = true;
			}
			if(!tsrTracking.getComplete().equals(complete)) {
				tsrTracking.setComplete(complete);
				isUpdate = true;
			}
			if(!tsrTracking.getWorkHours().equals(hours.doubleValue())) {
				tsrTracking.setWorkHours(hours.doubleValue());
				isUpdate = true;
			}
			if(!tsrTracking.getTotalTalkTime().equals(talkTime.doubleValue())) {
				tsrTracking.setTotalTalkTime(talkTime.doubleValue());
				isUpdate = true;
			}
			if(!tsrTracking.getNewUsed().equals(newUsed)) {
				tsrTracking.setNewUsed(newUsed);
				isUpdate = true;
			}
			if(!tsrTracking.getTotalPolicy().equals(totalPolicy)) {
				tsrTracking.setTotalPolicy(totalPolicy);
				isUpdate = true;
			}
			
			if(isUpdate) {
				tsrTrackingService.update(tsrTracking, LOGIN_USER);
			}
		} else {
			throw new Exception("Found tsr tracking more than 1 records b >> " + (tsr == null ? "tsrName: " + tsrName : "tsrCode: " + tsr.getTsrCode()) + " | trackingDate: " + trackingDate + " | Listlot: " + listLotCode);
		}
	}
	
	private Tsr getTsrByName(String tsrName) {
		if(StringUtils.isBlank(tsrName)) {logger.info("## TSR Name is Blank"); return null;}
		
		DetachedCriteria criteria = DetachedCriteria.forClass(Tsr.class);
//		criteria.add(Restrictions.eq("fullName", tsrName));
		criteria.add(Restrictions.like("fullName", tsrName, MatchMode.ANYWHERE));
		criteria.add(Restrictions.isNull("resignDate"));
		List<Tsr> list = null;
		
		try {
			list = tsrService.findByCriteria(criteria);
//			<!-- if not found, find without resign date is null -->
			if(list.isEmpty()) {
				criteria = DetachedCriteria.forClass(Tsr.class);
//				criteria.add(Restrictions.eq("fullName", tsrName));
				criteria.add(Restrictions.like("fullName", tsrName, MatchMode.ANYWHERE));
				list = tsrService.findByCriteria(criteria);
			}
			
			if(!list.isEmpty() && list.size() == 1) {
				return list.get(0);
			} else if(!list.isEmpty() && list.size() > 1) {
				Tsr temp = null;
				for(Tsr tsr : list) {
					if(temp == null) temp = tsr;
					
					if(temp.getEffectiveDate().compareTo(tsr.getEffectiveDate()) < 0) {
						temp = tsr;
					}
				}
				return temp;
			}
		} catch(Exception e) {
			logger.error(e.getMessage(), e);
		}
		
		return null;
	}
	
	private String removeTitle(String val) {
		if(val.startsWith(" ")) val = val.substring(1, val.length());
		for(String s : titles) {
			if(val.contains(s)) {
				return val.replace(s, "").trim();
			}
		}
		return val;
	}
	
	private String getListLotCode(String val) {
		String result = "";
		if(!StringUtils.isEmpty(val)) {
			int count = 0;
			
//			<!-- Check -->
			for(int i = 0; i < val.length(); i++) {
				if(val.charAt(i) == '(') {
					count++;
				}
			}
			
//			<!-- process -->
			if(count == 1) {
				return val.substring(val.indexOf("(") + 1, val.indexOf(")")).trim();
			} else if(count == 2) {
				return val.substring(val.indexOf("(", val.indexOf("(") + 1) + 1, val.length() - 1).trim();
			} else {
				logger.info("Cannot find Keycode");
			}
		}
		return result;
	}
	
	private String getFileFormatPath(String fileName) {
		if(fileName.contains("OTO")) {
			return EFileFormat.TSR_TRACKING_OTO.getValue();
		} else if(fileName.contains("TELE")) {
			return EFileFormat.TSR_TRACKING_TELE.getValue();
		}
		return "";
	}

	@Override
	public void importData(String path, String processDate) throws Exception {
		// TODO Auto-generated method stub
		
	}

}
