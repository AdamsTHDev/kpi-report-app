package com.adms.batch.kpireport.service.impl;

import java.io.FileInputStream;
import java.io.InputStream;
import java.util.Date;
import java.util.List;

import com.adms.batch.kpireport.enums.EFileFormat;
import com.adms.batch.kpireport.service.DataImporter;
import com.adms.batch.kpireport.util.AppConfig;
import com.adms.entity.Tsr;
import com.adms.entity.TsrHierarchy;
import com.adms.imex.excelformat.DataHolder;
import com.adms.imex.excelformat.ExcelFormat;
import com.adms.kpireport.service.TsrHierarchyService;
import com.adms.kpireport.service.TsrService;

public class SupDsmImporter implements DataImporter {
	
	private final String LOGIN_USER = "SUP_DSM_IMPORTER";

	private TsrHierarchyService tsrHierarchyService = (TsrHierarchyService) AppConfig.getInstance().getBean("tsrHierarchyService");
	private TsrService tsrService = (TsrService) AppConfig.getInstance().getBean("tsrService");
	
	@Override
	public void importData(final String path) throws Exception {
		
	}

	@Override
	public void importData(final String path, final String processDate) throws Exception {
		InputStream is = null;
		InputStream formatStream = null;
		
		try {
			is = new FileInputStream(path);
			formatStream = Thread.currentThread().getContextClassLoader().getResourceAsStream(EFileFormat.SUP_DSM.getValue());
			
			ExcelFormat ef = new ExcelFormat(formatStream);
			DataHolder wbHolder = ef.readExcel(is);
			DataHolder sheet = wbHolder.get(wbHolder.getKeyList().get(0));
			
//			<!-- Clear data by process date -->
			deleteHirarchyByDate(processDate);
			
			List<DataHolder> dataList = sheet.getDataList("dataList");
			for(DataHolder data : dataList) {
				String tsrCode = data.get("tsrCode").getStringValue();
				String uplineCode = data.get("uplineCode").getStringValue();
				Date startDate = (Date) data.get("startDate").getValue();
				Date endDate = (Date) data.get("endDate").getValue();
				
				Tsr tsr = tsrService.find(new Tsr(tsrCode)).get(0);
				Tsr upline = tsrService.find(new Tsr(uplineCode)).get(0);
				
				TsrHierarchy tsrHierarchy = new TsrHierarchy();
				tsrHierarchy.setTsr(tsr);
				tsrHierarchy.setUpline(upline);
				tsrHierarchy.setEffectiveDate(startDate);
				tsrHierarchy.setEndDate(endDate);
				
				tsrHierarchyService.add(tsrHierarchy, LOGIN_USER);
			}
			
		} catch(Exception e) {
			e.printStackTrace();
		} finally {
			try { is.close(); } catch(Exception e) {}
			try { formatStream.close(); } catch(Exception e) {}
		}
		
	}
	
	private void deleteHirarchyByDate(final String processDate) throws Exception {
		String hql = "from TsrHierarchy d "
				+ " where 1 = 1 "
				+ " and CONVERT(nvarchar(6), d.effectiveDate, 112) = ? "
				+ " and CONVERT(nvarchar(6), d.endDate, 112) = ? ";
		List<TsrHierarchy> list = this.tsrHierarchyService.findByHql(hql, processDate, processDate);
		for(TsrHierarchy t : list) {
			this.tsrHierarchyService.delete(t);
		}
		
	}

}
