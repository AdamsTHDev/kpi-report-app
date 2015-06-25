package com.adms.batch.kpireport.service.impl;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.math.BigDecimal;
import java.util.ArrayList;
import java.util.Calendar;
import java.util.Date;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

import org.apache.commons.lang3.StringUtils;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.FormulaEvaluator;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.ss.util.CellReference;
import org.hibernate.criterion.DetachedCriteria;
import org.hibernate.criterion.Order;
import org.hibernate.criterion.Projections;
import org.hibernate.criterion.Restrictions;

import com.adms.batch.kpireport.service.ReportExporter;
import com.adms.batch.kpireport.util.AppConfig;
import com.adms.batch.kpireport.util.FixMsigListLot;
import com.adms.entity.Campaign;
import com.adms.entity.KpiCategorySetup;
import com.adms.entity.KpiResult;
import com.adms.entity.Sales;
import com.adms.entity.Tsr;
import com.adms.entity.bean.KpiRetention;
import com.adms.kpireport.service.CampaignService;
import com.adms.kpireport.service.KpiCategorySetupService;
import com.adms.kpireport.service.KpiResultService;
import com.adms.kpireport.service.KpiRetentionService;
import com.adms.kpireport.service.SalesService;
import com.adms.kpireport.service.TsrService;
import com.adms.utils.DateUtil;
import com.adms.utils.FileUtil;
import com.adms.utils.Logger;

public class KpiReportExporter implements ReportExporter {

	private static Logger logger = Logger.getLogger();
	
	private final String templateKpiReportPath = "config/template/exportKpiTemplate.xlsx";

	private KpiCategorySetupService kpiCategorySetupService = (KpiCategorySetupService) AppConfig.getInstance().getBean("kpiCategorySetupService");
	
	private CampaignService campaignService = (CampaignService) AppConfig.getInstance().getBean("campaignService");
	
	private TsrService tsrService = (TsrService) AppConfig.getInstance().getBean("tsrService");
	
	private KpiResultService kpiResultService = (KpiResultService) AppConfig.getInstance().getBean("kpiResultService");
	
	private Map<String, List<String[]>> supGradeMap = new HashMap<>();
	
	private Map<String, Double[]> tsrABGradeByDsmMap = new HashMap<>();
	
	private String outPath = "";
	
	@Override
	public void exportExcel(final String destination, final String processDate) throws Exception {
		
		String yyyyMM = processDate.substring(0, 6);
		String msigListLots = "";
		outPath = destination + "/" + yyyyMM;
		
		try {
			msigListLots = FixMsigListLot.getInstance().getFixMisgListLotDelim(processDate);
			logger.info("MSIG WB List Lots >> " + msigListLots);

			logicExport(yyyyMM, msigListLots);
			
		} catch(Exception e) {
			logger.error(e.getMessage(), e);
		}
	}
	
	private void logicExport(String yyyyMM, String fixListLotCodeDelim) {
		logger.info("## Generating Report... ##");
		try {
			List<KpiResult> kpiResults = getKpiResultByYearMonth(yyyyMM);
			if(kpiResults.isEmpty()) return;
			
			List<KpiRetention> kpiRetentions = getKpiRetentions(yyyyMM);
			
			exportKpiReport(yyyyMM, fixListLotCodeDelim, kpiResults, kpiRetentions);
		} catch(Exception e) {
			logger.error(e.getMessage(), e);
		}
	}
	
	private void exportKpiReport(String yyyyMM, String fixListLotCodeDelim, List<KpiResult> kpiResults, List<KpiRetention> kpiRetentions) {
		List<String> dsmList = new ArrayList<>();
		Map<String, Map<String, Map<String, List<String>>>> mapCampaignDsm = new HashMap<>();
		String currentCampaignCode = "";
		
		try {
			/**
			 * loop for retrieving data tree
			 */
			for(KpiResult kpiResult : kpiResults) {
				
//				<!-- for All DSM Code list -->
				if(kpiResult.getDsm() == null) throw new Exception("DSM not found ==> campaignCode: " + kpiResult.getCampaign().getCampaignCode());
				
				if(!dsmList.contains(kpiResult.getDsm().getTsrCode())) {
					dsmList.add(kpiResult.getDsm().getTsrCode());
				}

//				<!-- since MSIG Broker POM WB is separated by List lot code -->
				currentCampaignCode = isMsigWB(fixListLotCodeDelim, kpiResult.getListLotCode()) ? kpiResult.getCampaign().getCampaignCode() + "_" + kpiResult.getListLotCode() : kpiResult.getCampaign().getCampaignCode();
				
				if(mapCampaignDsm.get(currentCampaignCode) == null) {
					mapCampaignDsm.put(currentCampaignCode, new HashMap<String, Map<String, List<String>>>());
				}

				if(mapCampaignDsm.get(currentCampaignCode).get(kpiResult.getDsm().getTsrCode()) == null) {
					mapCampaignDsm.get(currentCampaignCode).put(kpiResult.getDsm().getTsrCode(), new HashMap<String, List<String>>());
				}
				
				if(mapCampaignDsm.get(currentCampaignCode).get(kpiResult.getDsm().getTsrCode()).get(kpiResult.getTsm().getTsrCode()) == null) {
					mapCampaignDsm.get(currentCampaignCode).get(kpiResult.getDsm().getTsrCode()).put(kpiResult.getTsm().getTsrCode(), new ArrayList<String>());
				}
				
				mapCampaignDsm.get(currentCampaignCode).get(kpiResult.getDsm().getTsrCode()).get(kpiResult.getTsm().getTsrCode()).add(kpiResult.getTsr().getTsrCode());
			}
			
//			<!-- Process -->
			processKpiReportByCampaign(yyyyMM, mapCampaignDsm, kpiRetentions);
			
			processKpiReportForDSM(yyyyMM, dsmList, kpiRetentions);
			
			processKpiReportForSupGradeSummary(yyyyMM, fixListLotCodeDelim);
			
		} catch(Exception e) {
			logger.error(e.getMessage(), e);
		}
	}
	
	private void processKpiReportForSupGradeSummary(String yearMonth, String fixListLotCodeDelim) {
		InputStream templateStream = null;
		Workbook wb = null;
		Sheet tempSheet = null;
		Sheet sheet = null;
		int numOfTemplates = 0;
		
		try {
//			<!-- initial workbook and sheet -->
			String sheetName = "SUPs Grade Summary";
			
//			<!-- get template Workbook -->
			templateStream = getTemplateStream();
			wb = WorkbookFactory.create(templateStream);
			
//			<!-- Template location -->
			tempSheet = wb.getSheetAt(2);
			numOfTemplates = wb.getNumberOfSheets();
			
//			<!-- New sheet for report -->
			sheet = wb.createSheet(sheetName);
			
//			<!-- process table header -->
			copyRowCellDataWithSameColumn(tempSheet, sheet, 0, 0, getColumnIndex("A"), getColumnIndex("D"), true);
			
//			<!-- process data -->
			for(String key : supGradeMap.keySet()) {
				List<String[]> vals = supGradeMap.get(key);
				
				logicSupGradeSummary(tempSheet, sheet, key, vals);
			}
			
//			<!-- copy columns width -->
			for(int n = 0; n < 4; n++) sheet.setColumnWidth(n, tempSheet.getColumnWidth(n));
			
//			set Grid blank
			sheet.setDisplayGridlines(false);
			
			writeout(wb, outPath, "SUPs Grade Summary_" + yearMonth +  ".xlsx", numOfTemplates);
			
		} catch(Exception e) {
			logger.error(e.getMessage(), e);
		}
	}
	
	private void processKpiReportByCampaign(String yearMonth, Map<String, Map<String, Map<String, List<String>>>> mapCampaignDsm, List<KpiRetention> kpiRetentions) {
		InputStream templateStream = null;
		Workbook wb = null;
		Sheet tempSheet = null;
		Sheet sheet = null;
		Map<String, Double[]> retentionForSup = null;
		int numOfTemplates = 0;
		
		try {
			int totalWorkDays = getWorkDays(yearMonth);
			logger.info("# Total Working day: " + totalWorkDays);
			
			retentionForSup = new HashMap<>();
			for(KpiRetention kpiRetention : kpiRetentions) {
				sumRetention(retentionForSup, kpiRetention.getSupCode()
						, kpiRetention.getBeginMonth().doubleValue()
						, kpiRetention.getDuringMonth().doubleValue()
						, kpiRetention.getEndMonth().doubleValue());
//				if(retentionForSup.get(kpiRetention.getSupCode()) == null) {
//					retentionForSup.put(kpiRetention.getSupCode(), new Double[]{0D, 0D, 0D});
//				}
//				Double[] retentions = retentionForSup.get(kpiRetention.getSupCode());
//				retentions[0] += kpiRetention.getBeginMonth();
//				retentions[1] += kpiRetention.getDuringMonth();
//				retentions[2] += kpiRetention.getEndMonth();
			}
			
			for(String campaignCode : mapCampaignDsm.keySet()) {
				
//				<!-- initial workbook and sheet -->
				String sheetName = campaignCode.contains("_") ? "MSIG Broker POM WB OTO" : campaignService.find(new Campaign(campaignCode)).get(0).getCampaignNameMgl();
				
//				<!-- get template Workbook -->
				templateStream = getTemplateStream();
				wb = WorkbookFactory.create(templateStream);
				
//				<!-- Template location -->
				tempSheet = wb.getSheetAt(0);
				numOfTemplates = wb.getNumberOfSheets();
				
//				<!-- New sheet for report -->
				sheet = wb.createSheet(sheetName);
				
//				<!-- process table header -->
				copyRowCellDataWithSameColumn(tempSheet, sheet, 0, 0, getColumnIndex("A"), getColumnIndex("G"), true);
				
//				<!-- Report processing -->
				for(String dsmCode : mapCampaignDsm.get(campaignCode).keySet()) {
//					<!-- Skip. will be processed later -->
//					<!-- DSM Section -->
//					logger.info("--| " + dsmCode);
					
					for(String supCode : mapCampaignDsm.get(campaignCode).get(dsmCode).keySet()) {
//						<!-- SUP Section -->
						Object[] supObjects = getSupKpiByCampaign(yearMonth, campaignCode, dsmCode, supCode);
						logicSupSectionByCampaign(tempSheet, sheet, yearMonth, supObjects, retentionForSup);
						
						for(String tsrCode : mapCampaignDsm.get(campaignCode).get(dsmCode).get(supCode)) {
//							<!-- TSR Section -->
							Object[] tsrObjects = getTsrKpiByCampaign(yearMonth, campaignCode, dsmCode, supCode, tsrCode);
							logicTsrSectionByCampaign(tempSheet, sheet, yearMonth, tsrObjects, totalWorkDays);
						}
					}
				}
				
//				<!-- Grade Information -->
				copyRowCellDataWithSameColumn(tempSheet, sheet, 0, 0, getColumnIndex("J"), getColumnIndex("K"), true);
				copyRowCellDataWithSameColumn(tempSheet, sheet, 1, 1, getColumnIndex("J"), getColumnIndex("K"), true);
				copyRowCellDataWithSameColumn(tempSheet, sheet, 2, 2, getColumnIndex("J"), getColumnIndex("K"), true);
				copyRowCellDataWithSameColumn(tempSheet, sheet, 3, 3, getColumnIndex("J"), getColumnIndex("K"), true);

//				<!-- copy columns width -->
				for(int n = 0; n < 11; n++) sheet.setColumnWidth(n, tempSheet.getColumnWidth(n));
				
//				set Grid blank
				sheet.setDisplayGridlines(false);
				
				writeout(wb, outPath, sheetName + "_" + yearMonth + ".xlsx", numOfTemplates);
			}
			
			
		} catch (InvalidFormatException | IOException e) {
			logger.error(e.getMessage(), e);
		} catch (Exception e) {
			logger.error(e.getMessage(), e);
		} finally {
			try {wb.close();} catch(Exception e) {}
			try {templateStream.close();} catch(Exception e) {}
		}
	}
	
	private void processKpiReportForDSM(String yearMonth, List<String> dsmList, List<KpiRetention> kpiRetentions) {
		InputStream templateStream = null;
		Workbook wb = null;
		Sheet tempSheet = null;
		Sheet sheet = null;
		Map<String, Double[]> retentionForDsm = new HashMap<>();
		int numOfTemplates = 0;
		
		try {
			for(KpiRetention kpiRetention : kpiRetentions) {
				sumRetention(retentionForDsm, kpiRetention.getDsmCode()
						, kpiRetention.getBeginMonth().doubleValue()
						, kpiRetention.getDuringMonth().doubleValue()
						, kpiRetention.getEndMonth().doubleValue());
			}
			
//			<!-- initial workbook and sheet -->
			String sheetName = "DSMs KPI";
			
//			<!-- get template Workbook -->
			templateStream = getTemplateStream();
			wb = WorkbookFactory.create(templateStream);
			
//			<!-- Template location -->
			tempSheet = wb.getSheetAt(1);
			numOfTemplates = wb.getNumberOfSheets();
			
//			<!-- New sheet for report -->
			sheet = wb.createSheet(sheetName);
			
//			<!-- process table header -->
			copyRowCellDataWithSameColumn(tempSheet, sheet, 0, 0, getColumnIndex("A"), getColumnIndex("G"), true);

//			<!-- Report Processing -->
			for(String dsmCode : dsmList) {
//				<!-- getting DSM KPI -->
				Object[] objects = getDsmKpi(yearMonth, dsmCode);
				logicDsmSection(tempSheet, sheet, yearMonth, objects, retentionForDsm);
			}

//			<!-- copy columns width -->
			for(int n = 0; n < 8; n++) sheet.setColumnWidth(n, tempSheet.getColumnWidth(n));
			
//			set Grid blank
			sheet.setDisplayGridlines(false);
			
			writeout(wb, outPath, sheetName + "_" + yearMonth + ".xlsx", numOfTemplates);
			
		} catch(Exception e) {
			logger.error(e.getMessage(), e);
		} finally {
			try {templateStream.close();} catch(Exception e) {}
			try {wb.close();} catch(Exception e) {};
		}
	}

	private void logicDsmSection(Sheet tempSheet, Sheet sheet, String yearMonth, Object[] objects, Map<String, Double[]> retentionForDsm) throws Exception {
		/*
		 * object index
		 * [0] = dsmCode
		 * [1] = totalAfyp
		 * [2] = firstConfirmSale
		 * [3] = allSale
		 */
		String dsmCode = String.valueOf(objects[0]);
		BigDecimal totalAfyp = new BigDecimal(String.valueOf(objects[1]));
		Integer firstConfirmSale = Integer.valueOf(String.valueOf(objects[2]));
		Integer allSale = Integer.valueOf(String.valueOf(objects[3]));
	
		String fullName = tsrService.find(new Tsr(dsmCode)).get(0).getFullName();
		
		boolean skip = false;
		
//		<!-- get end of month -->
		Calendar end = DateUtil.getCurrentCalendar();
		end.setTime(DateUtil.convStringToDate("yyyyMMdd", yearMonth + "01"));
		DateUtil.addMonth(end, 1);
		DateUtil.addDay(end, -1);
		
//		<!-- Get Kpi Category -->
		List<KpiCategorySetup> kpiCategories = getKpiCategory(dsmCode, "DSM", null, null, DateUtil.convStringToDate("yyyyMMdd", yearMonth + "01"), end.getTime());
		if(kpiCategories.isEmpty()) {
			logger.error("No KPI Category for this DSM: " + dsmCode + " effectiveDate: " + DateUtil.convStringToDate("yyyyMMdd", yearMonth + "01") + " | endDate: " + end.getTime());
			skip = true;
		}
		
//		<!-- DSM Template start on 2nd row and end at 6th row -->
		for(int rt = 1; rt < 6; rt++) {
			int currentRow = sheet.getLastRowNum() + 1;
			Row row = sheet.createRow(currentRow);
			
			for(int c = getColumnIndex("A"); c <= getColumnIndex("G"); c++) {
				boolean copyData = false;
				if((c == getColumnIndex("A") && rt != 2) || c == getColumnIndex("B")) copyData = true;
				copyRowCellDataWithSameColumn(tempSheet, sheet, rt, currentRow, c, c, copyData);
				
				if(!copyData && !skip) {
					Cell cell = row.getCell(c, Row.CREATE_NULL_AS_BLANK);
					/*
					 * switch Column
					 * case 0 = Position
					 * case 1 = KPIs
					 * case 2 = Weight
					 * case 3 = Target
					 * case 4 = Actual
					 * case 5 = %vs. Target
					 * case 6 = Score
					 */
					if(rt == 1) {
						String formula = "";
						switch(c) {
						case 2 : cell.setCellValue(kpiCategories.get(0).getWeight().doubleValue()); break;
						case 3 : cell.setCellValue(kpiCategories.get(0).getTarget().doubleValue()); break;
						case 4 : cell.setCellValue(totalAfyp.doubleValue()); break;
						case 5 :
							formula = "IF(D#ROW > 0, E#ROW / D#ROW, 0)";
							cell.setCellFormula(formula.replaceAll("#ROW", String.valueOf(currentRow + 1)));
							break;
						case 6 :
							formula = "IF(F#ROW > 1, C#ROW, F#ROW * C#ROW)";
							cell.setCellFormula(formula.replaceAll("#ROW", String.valueOf((currentRow + 1))));
							break;
						default : break;
						}
					} else if(rt == 2) {
						String formula = "";
						switch(c) {
						case 0 : cell.setCellValue(fullName); break;
						case 2 : cell.setCellValue(kpiCategories.get(1).getWeight().doubleValue()); break;
						case 3 : cell.setCellValue(kpiCategories.get(1).getTarget().doubleValue()); break;
						case 4 : cell.setCellValue(firstConfirmSale.doubleValue() / allSale.doubleValue()); break;
						case 5 :
							formula = "IF(D#ROW > 0, E#ROW / D#ROW, 0)";
							cell.setCellFormula(formula.replaceAll("#ROW", String.valueOf(currentRow + 1)));
							break;
						case 6 :
							formula = "IF(F#ROW > 1, C#ROW, F#ROW * C#ROW)";
							cell.setCellFormula(formula.replaceAll("#ROW", String.valueOf((currentRow + 1))));
							break;
						default : break;
						}
					} else if(rt == 3) {
						String formula = "";
						switch(c) {
						case 2 : cell.setCellValue(kpiCategories.get(2).getWeight().doubleValue()); break;
						case 3 : cell.setCellValue(kpiCategories.get(2).getTarget().doubleValue()); break;
						case 4 : 
							Double[] retention = retentionForDsm.get(dsmCode);
							Double base = retention[0].doubleValue() + retention[1].doubleValue();
							cell.setCellValue(base.compareTo(new Double(0d)) <= 0 ? 0d : retention[2].doubleValue() / base.doubleValue());
							break;
						case 5 :
							formula = "IF(D#ROW > 0, E#ROW / D#ROW, 0)";
							cell.setCellFormula(formula.replaceAll("#ROW", String.valueOf(currentRow + 1)));
							break;
						case 6 : 
							formula = "IF(F#ROW > 1, C#ROW, F#ROW * C#ROW)";
							cell.setCellFormula(formula.replaceAll("#ROW", String.valueOf((currentRow + 1))));
							break;
						default : break;
						}
					} else if(rt == 4) {
						String formula = "";
						switch(c) {
						case 2 : cell.setCellValue(kpiCategories.get(3).getWeight().doubleValue()); break;
						case 3 : cell.setCellValue(kpiCategories.get(3).getTarget().doubleValue()); break;
						case 4 : 
							Double[] abCat = this.tsrABGradeByDsmMap.get(dsmCode);
							cell.setCellValue(abCat[0] / abCat[1]); 
							break;
						case 5 :
							formula = "IF(D#ROW > 0, E#ROW / D#ROW, 0)";
							cell.setCellFormula(formula.replaceAll("#ROW", String.valueOf(currentRow + 1)));
							break;
						case 6 : 
							formula = "IF(F#ROW > 1, C#ROW, F#ROW * C#ROW)";
							cell.setCellFormula(formula.replaceAll("#ROW", String.valueOf((currentRow+1))));
							break;
						default : break;
						}
					} else if(rt == 5) {
						String sumFormula = "";
						switch(c) {
						case 2 : 
							sumFormula = "SUM(C" + (cell.getRowIndex() - 3) + ":C" + (cell.getRowIndex()) + ")";
							cell.setCellFormula(sumFormula);
							break;
						case 6 : 
							sumFormula = "SUM(G" + (cell.getRowIndex() - 3) + ":G" + (cell.getRowIndex()) + ")";
							cell.setCellFormula(sumFormula); 
							break;
						default : break;
						}
					}
				}
			}
		}
		
//		<!-- Calculate Grade for DSM -->
		Cell sumScoreCell = sheet.getRow(sheet.getLastRowNum()).getCell(getColumnIndex("G"), Row.CREATE_NULL_AS_BLANK);
		Cell gradeCell = sheet.getRow(sheet.getLastRowNum() - 4).createCell(getColumnIndex("H"), Cell.CELL_TYPE_STRING);
		FormulaEvaluator evaluator = sheet.getWorkbook().getCreationHelper().createFormulaEvaluator();
		Double score = skip ? null : evaluator.evaluate(sumScoreCell).getNumberValue();
		gradeCell.setCellValue(getGrade(score));
		
	}
	
	private void logicSupGradeSummary(Sheet tempSheet, Sheet sheet, String supCode, List<String[]> vals) throws Exception {
			String fullName = tsrService.find(new Tsr(supCode)).get(0).getFullName();
			int valSize = vals.size();
			int tempRow = 1;
			Double sumScore = 0D;
			int countCampaign = 0;
			int currentRow = 0;
			
			boolean isCopy = false;
			
			for(String[] val : vals) {
				
				currentRow = sheet.getLastRowNum() + 1;
				Row row = sheet.createRow(currentRow);
				
				for(int c = getColumnIndex("A"); c <= getColumnIndex("D"); c++) {
					if((c == getColumnIndex("A") && tempRow == 1) || (c == getColumnIndex("C") && tempRow == 4)) isCopy = true;
					copyRowCellDataWithSameColumn(tempSheet, sheet, tempRow, currentRow, c, c, isCopy);
				}
				
				String campaignCode = val[0].contains("_") ? val[0].substring(0, val[0].indexOf("_")) : val[0];
				String campaignName = val[0].contains("_") ? "MSIG Broker POM WB_OTO" : campaignService.find(new Campaign(campaignCode)).get(0).getCampaignNameMgl();
				Double score = Double.valueOf(val[1]);
				
				if(tempRow == 2) {
					row.getCell(0, Row.CREATE_NULL_AS_BLANK).setCellValue(fullName);
				}
	
				if(tempRow == 4) {
					copyRowCellDataWithSameColumn(tempSheet, sheet, tempRow - 1, currentRow, getColumnIndex("A"), getColumnIndex("D"), false);
				}
	
				row.getCell(1, Row.CREATE_NULL_AS_BLANK).setCellValue(campaignCode);
				row.getCell(2, Row.CREATE_NULL_AS_BLANK).setCellValue(campaignName);
				row.getCell(3, Row.CREATE_NULL_AS_BLANK).setCellValue(getGrade(score));
	
				sumScore+=score;
				countCampaign++;
				
				if(tempRow < 4) tempRow++;
			}
	
			int loopForBlank = 3;
			if(valSize < 3) {
				for(int n = 0; n < (loopForBlank - valSize); n++) {
					currentRow = sheet.getLastRowNum() + 1;
					copyRowCellDataWithSameColumn(tempSheet, sheet, 3, currentRow, getColumnIndex("A"), getColumnIndex("D"), false);
					if(tempRow == 2) {
						sheet.getRow(currentRow).getCell(0, Row.CREATE_NULL_AS_BLANK).setCellValue(fullName);
					}
					if(tempRow < 4) tempRow++;
				}
	
			}
			
	//		<!-- Summary Row -->
			currentRow = sheet.getLastRowNum() + 1;
			copyRowCellDataWithSameColumn(tempSheet, sheet, 4, currentRow, getColumnIndex("A"), getColumnIndex("D"), false);
			copyRowCellDataWithSameColumn(tempSheet, sheet, 4, currentRow, getColumnIndex("C"), getColumnIndex("C"), true);
			Double avgScore = sumScore / countCampaign;
			sheet.getRow(currentRow).getCell(getColumnIndex("D"), Row.CREATE_NULL_AS_BLANK).setCellValue(getGrade(avgScore));
		}

	private void logicSupSectionByCampaign(Sheet tempSheet, Sheet sheet, String yearMonth, Object[] objs, Map<String, Double[]> retentionForSup) throws Exception {
			/*
			 * object index 
			 * [0] = campaignCode
			 * [1] = listLotCode
			 * [2] = dsmCode
			 * [3] = supCode
			 * [4] = totalAfyp
			 * [5] = countTsr
			 * [6] = firstConfirmSale
			 * [7] = allSale
			 * [8] = successPolicy
			 * [9] = totalUsed
			 */
			String campaignCode = String.valueOf(objs[0]);
			String listLotCode = objs[1] == null ? null : String.valueOf(objs[1]);
	//		String dsmCode = String.valueOf(objs[2]);
			String supCode = String.valueOf(objs[3]);
			BigDecimal totalAfyp = new BigDecimal(String.valueOf(objs[4]));
			Integer countTsr = Integer.valueOf(String.valueOf(objs[5]));
			Integer firstConfirmSale = Integer.valueOf(String.valueOf(objs[6]));
			Integer allSale = Integer.valueOf(String.valueOf(objs[7]));
			Integer successPolicy = Integer.valueOf(String.valueOf(objs[8]));
			Integer totalUsed = Integer.valueOf(String.valueOf(objs[9]));
			
			boolean skip = false;
			
			String fullName = tsrService.find(new Tsr(supCode)).get(0).getFullName();
			
	//		<!-- get end of month -->
			Calendar end = DateUtil.getCurrentCalendar();
			end.setTime(DateUtil.convStringToDate("yyyyMMdd", yearMonth + "01"));
			DateUtil.addMonth(end, 1);
			DateUtil.addDay(end, -1);
			
			List<KpiCategorySetup> kpiCategories = getKpiCategory(supCode, "SUP", campaignCode, listLotCode, DateUtil.convStringToDate("yyyyMMdd", yearMonth + "01"), end.getTime());
			if(kpiCategories.isEmpty()) {
				logger.warn("No KPI Category for this SUP: " + supCode + " | campaignCode: " + campaignCode + " | listLotCode: " + listLotCode + " effectiveDate: " + DateUtil.convStringToDate("yyyyMMdd", yearMonth + "01") + " | endDate: " + end.getTime());
				skip = true;
			}
			
			
	//		<!-- Sup template start on 2nd row and end on 7th row -->
			for(int rt = 1; rt < 7; rt++) {
				int currentRow = sheet.getLastRowNum() + 1;
				Row row = sheet.createRow(currentRow);
				
				for(int c = getColumnIndex("A"); c <= getColumnIndex("G"); c++) {
					boolean copyData = false;
					if((c == getColumnIndex("A") && rt != 2) || c == getColumnIndex("B")) copyData = true;
					copyRowCellDataWithSameColumn(tempSheet, sheet, rt, currentRow, c, c, copyData);
					
					if(!copyData && !skip) {
						Cell cell = row.getCell(c, Row.CREATE_NULL_AS_BLANK);
						/*
						 * switch Column
						 * case 0 = Position
						 * case 1 = KPIs
						 * case 2 = Weight
						 * case 3 = Target
						 * case 4 = Actual
						 * case 5 = %vs. Target
						 * case 6 = Score
						 */
						if(rt == 1) {
							String formula = "";
							switch(c) {
							case 2 : 
								cell.setCellValue(kpiCategories.get(0).getWeight().doubleValue()); 
								cell.getSheet().addMergedRegion(new CellRangeAddress(cell.getRowIndex(), cell.getRowIndex() + 1, cell.getColumnIndex(), cell.getColumnIndex()));
								break;
							case 3 : cell.setCellValue(kpiCategories.get(0).getTarget().doubleValue()); break;
							case 4 : cell.setCellValue(totalAfyp.doubleValue() / countTsr.doubleValue()); break;
							case 6 :
								formula = "IF(F#ROW > 1, C#ROW, F#ROW * C#ROW)";
								cell.setCellFormula(formula.replaceAll("#ROW", String.valueOf((currentRow+1))));
								break;
							default : break;
							}
						} else if(rt == 2) {
							String formula = "";
							switch(c) {
							case 0 : cell.setCellValue(fullName); break;
							case 3 : cell.setCellValue(kpiCategories.get(1).getTarget().doubleValue()); break;
							case 4 : cell.setCellValue(totalAfyp.doubleValue()); break;
							case 5 : 
								formula = "(E" + (cell.getRowIndex()) + "/" + "D" + (cell.getRowIndex()) + ")*(E" + (cell.getRowIndex() + 1) + "/D" + (cell.getRowIndex() + 1) + ")";
								cell.getSheet().addMergedRegion(new CellRangeAddress(cell.getRowIndex() - 1, cell.getRowIndex(), cell.getColumnIndex(), cell.getColumnIndex()));
								Cell mergedCell = sheet.getRow(cell.getRowIndex() - 1).getCell(c, Row.CREATE_NULL_AS_BLANK);
								mergedCell.setCellFormula(formula);
								break;
							case 6 : 
								cell.getSheet().addMergedRegion(new CellRangeAddress(cell.getRowIndex() - 1, cell.getRowIndex(), cell.getColumnIndex(), cell.getColumnIndex()));
								break;
							default : break;
							}
						} else if(rt == 3) {
							String formula = "";
							switch(c) {
							case 2 : cell.setCellValue(kpiCategories.get(2).getWeight().doubleValue()); break;
							case 3 : cell.setCellValue(kpiCategories.get(2).getTarget().doubleValue()); break;
							case 4 : cell.setCellValue(totalUsed.doubleValue() == 0d ? 0D : successPolicy.doubleValue() / totalUsed.doubleValue()); break;
							case 5 : 
								formula = "IF(D#ROW > 0, E#ROW / D#ROW, 0)";
								cell.setCellFormula(formula.replaceAll("#ROW", String.valueOf((currentRow+1))));
								break;
							case 6 : 
								formula = "IF(F#ROW > 1, C#ROW, F#ROW * C#ROW)";
								cell.setCellFormula(formula.replaceAll("#ROW", String.valueOf((currentRow+1))));
								break;
							default : break;
							}
						} else if(rt == 4) {
							String formula = "";
							switch(c) {
							case 2 : cell.setCellValue(kpiCategories.get(3).getWeight().doubleValue()); break;
							case 3 : cell.setCellValue(kpiCategories.get(3).getTarget().doubleValue()); break;
							case 4 : cell.setCellValue(firstConfirmSale.doubleValue() / allSale.doubleValue()); break;
							case 5 : 
								formula = "IF(D#ROW > 0, E#ROW / D#ROW, 0)";
								cell.setCellFormula(formula.replaceAll("#ROW", String.valueOf((currentRow+1))));
								break;
							case 6 : 
								formula = "IF(F#ROW > 1, C#ROW, F#ROW * C#ROW)";
								cell.setCellFormula(formula.replaceAll("#ROW", String.valueOf((currentRow+1))));
								break;
							default : break;
							}
						} else if(rt == 5) {
							String formula = "";
							switch(c) {
							case 2 : cell.setCellValue(kpiCategories.get(4).getWeight().doubleValue()); break;
							case 3 : cell.setCellValue(kpiCategories.get(4).getTarget().doubleValue()); break;
							case 4 : 
								Double[] retention = retentionForSup.get(supCode);
								Double base = retention[0].doubleValue() + retention[1].doubleValue();
								cell.setCellValue(base.compareTo(new Double(0d)) <= 0 ? 0d : retention[2].doubleValue() / base.doubleValue());
								break;
							case 5 : 
								formula = "IF(D#ROW > 0, E#ROW / D#ROW, 0)";
								cell.setCellFormula(formula.replaceAll("#ROW", String.valueOf((currentRow+1))));
								break;
							case 6 : 
								formula = "IF(F#ROW > 1, C#ROW, F#ROW * C#ROW)";
								cell.setCellFormula(formula.replaceAll("#ROW", String.valueOf((currentRow+1))));
								break;
							default : break;
							}
						} else if(rt == 6) {
							String sumFormula = "";
							switch(c) {
							case 2 : 
								sumFormula = "SUM(C" + (cell.getRowIndex() - 4) + ":C" + (cell.getRowIndex()) + ")";
								cell.setCellFormula(sumFormula);
								break;
							case 6 : 
								sumFormula = "SUM(G" + (cell.getRowIndex() - 4) + ":G" + (cell.getRowIndex()) + ")";
								cell.setCellFormula(sumFormula); 
								break;
							default : break;
							}
						}
					} 
					
				}
				
	//			<!-- set SUP name -->
				if(skip && rt == 2) {
					row.getCell(0, Row.CREATE_NULL_AS_BLANK).setCellValue(fullName);
				}
				
			}
			
	//		<!-- Calculate Grade for SUP -->
			Cell sumScoreCell = sheet.getRow(sheet.getLastRowNum()).getCell(getColumnIndex("G"), Row.CREATE_NULL_AS_BLANK);
			Cell gradeCell = sheet.getRow(sheet.getLastRowNum() - 5).createCell(getColumnIndex("H"), Cell.CELL_TYPE_STRING);
			FormulaEvaluator evaluator = sheet.getWorkbook().getCreationHelper().createFormulaEvaluator();
			Double score = skip ? null : evaluator.evaluate(sumScoreCell).getNumberValue();
			gradeCell.setCellValue(getGrade(score));
			
	//		<!-- Keep score of sup in map. For Sup grade summary -->
			if(!skip) {
				if(supGradeMap.get(supCode) == null) {
					supGradeMap.put(supCode, new ArrayList<String[]>());
				}
				supGradeMap.get(supCode).add(new String[]{campaignCode.concat(StringUtils.isNoneBlank(listLotCode) ? ("_" + listLotCode) : ""), score.toString()});
			}
		}

	private void logicTsrSectionByCampaign(Sheet tempSheet, Sheet sheet, String yearMonth, Object[] objs, Integer totalWorkDays) throws Exception {
		/*
		 * object index
		 * [0] = campaignCode
		 * [1] = listLotCode
		 * [2] = dsmCode
		 * [3] = supCode
		 * [4] = tsrCode
		 * [5] = totalAfyp
		 * [6] = attendance
		 * [7] = totalTalkHrs
		 * [8] = firstConfirmSale
		 * [9] = allSale
		 */
		String campaignCode = String.valueOf(objs[0]);
		String listLotCode = objs[1] == null ? null : String.valueOf(objs[1]);
		String dsmCode = String.valueOf(objs[2]);
//		String supCode = String.valueOf(objs[3]);
		String tsrCode = String.valueOf(objs[4]);
		BigDecimal totalAfyp = new BigDecimal(String.valueOf(objs[5]));
		Integer attendance = Integer.valueOf(String.valueOf(objs[6]));
		BigDecimal totalTalkHrs = new BigDecimal(String.valueOf(objs[7]));
		Integer firstConfirmSale = Integer.valueOf(String.valueOf(objs[8]));
		Integer allSale = Integer.valueOf(String.valueOf(objs[9]));
		
		boolean skip = false;
		
		String fullName = tsrService.find(new Tsr(tsrCode)).get(0).getFullName();
		
//		<!-- get end of month -->
		Calendar end = DateUtil.getCurrentCalendar();
		end.setTime(DateUtil.convStringToDate("yyyyMMdd", yearMonth + "01"));
		DateUtil.addMonth(end, 1);
		DateUtil.addDay(end, -1);
		
		List<KpiCategorySetup> kpiCategories = getKpiCategory(null, "TSR", campaignCode, listLotCode, DateUtil.convStringToDate("yyyyMMdd", yearMonth + "01"), end.getTime());
		if(kpiCategories.isEmpty()) {
			logger.warn("No KPI Category for this TSR: " + tsrCode + " | campaignCode: " + campaignCode + " | listLotCode: " + listLotCode + " effectiveDate: " + DateUtil.convStringToDate("yyyyMMdd", yearMonth + "01") + " | endDate: " + end.getTime());
			skip = true;
		}
		
//		<!-- TSR template start on 8nd row and end on 12th row -->
		for(int rt = 7; rt < 12; rt++) {
			int currentRow = sheet.getLastRowNum() + 1;
			Row row = sheet.createRow(currentRow);
			
			for(int c = getColumnIndex("A"); c <= getColumnIndex("G"); c++) {
				boolean copyData = false;
				if((c == getColumnIndex("A") && rt != 8) || c == getColumnIndex("B")) copyData = true;
				copyRowCellDataWithSameColumn(tempSheet, sheet, rt, currentRow, c, c, copyData);
				
				if(!copyData && !skip) {
					Cell cell = row.getCell(c, Row.CREATE_NULL_AS_BLANK);
					/*
					 * switch Column
					 * case 0 = Position
					 * case 1 = KPIs
					 * case 2 = Weight
					 * case 3 = Target
					 * case 4 = Actual
					 * case 5 = %vs. Target
					 * case 6 = Score
					 */
					if(rt == 7) {
						String formula = "";
						switch(c) {
						case 2 : cell.setCellValue(kpiCategories.get(0).getWeight().doubleValue()); break;
						case 3 : cell.setCellValue(kpiCategories.get(0).getTarget().doubleValue()); break;
						case 4 : cell.setCellValue(totalAfyp.doubleValue()); break;
						case 5 :
							formula = "IF(D#ROW > 0, E#ROW / D#ROW, 0)";
							cell.setCellFormula(formula.replaceAll("#ROW", String.valueOf(currentRow + 1)));
							break;
						case 6 :
							formula = "IF(F#ROW > 1, C#ROW, F#ROW * C#ROW)";
							cell.setCellFormula(formula.replaceAll("#ROW", String.valueOf((currentRow + 1))));
							break;
						default : break;
						}
					} else if(rt == 8) {
						String formula = "";
						switch(c) {
						case 0 : cell.setCellValue(fullName); break;
						case 2 : cell.setCellValue(kpiCategories.get(1).getWeight().doubleValue()); break;
						case 3 : cell.setCellValue(kpiCategories.get(1).getTarget().doubleValue()); break;
						case 4 : cell.setCellValue(attendance.doubleValue() / totalWorkDays.doubleValue()); break;
						case 5 :
							formula = "IF(D#ROW > 0, E#ROW / D#ROW, 0)";
							cell.setCellFormula(formula.replaceAll("#ROW", String.valueOf(currentRow + 1)));
							break;
						case 6 :
							formula = "IF(F#ROW > 1, C#ROW, F#ROW * C#ROW)";
							cell.setCellFormula(formula.replaceAll("#ROW", String.valueOf((currentRow + 1))));
							break;
						default : break;
						}
					} else if(rt == 9) {
						String formula = "";
						switch(c) {
						case 2 : cell.setCellValue(kpiCategories.get(2).getWeight().doubleValue()); break;
						case 3 : cell.setCellValue(kpiCategories.get(2).getTarget().doubleValue()); break;
						case 4 : cell.setCellValue(totalTalkHrs.doubleValue()); break;
						case 5 :
							formula = "IF(D#ROW > 0, E#ROW / D#ROW, 0)";
							cell.setCellFormula(formula.replaceAll("#ROW", String.valueOf(currentRow + 1)));
							break;
						case 6 : 
							formula = "IF(F#ROW > 1, C#ROW, F#ROW * C#ROW)";
							cell.setCellFormula(formula.replaceAll("#ROW", String.valueOf((currentRow + 1))));
							break;
						default : break;
						}
					} else if(rt == 10) {
						String formula = "";
						switch(c) {
						case 2 : cell.setCellValue(kpiCategories.get(3).getWeight().doubleValue()); break;
						case 3 : cell.setCellValue(kpiCategories.get(3).getTarget().doubleValue()); break;
						case 4 : cell.setCellValue(firstConfirmSale.doubleValue() / allSale.doubleValue()); break;
						case 5 :
							formula = "IF(D#ROW > 0, E#ROW / D#ROW, 0)";
							cell.setCellFormula(formula.replaceAll("#ROW", String.valueOf(currentRow + 1)));
							break;
						case 6 : 
							formula = "IF(F#ROW > 1, C#ROW, F#ROW * C#ROW)";
							cell.setCellFormula(formula.replaceAll("#ROW", String.valueOf((currentRow+1))));
							break;
						default : break;
						}
					} else if(rt == 11) {
						String sumFormula = "";
						switch(c) {
						case 2 : 
							sumFormula = "SUM(C" + (cell.getRowIndex() - 3) + ":C" + (cell.getRowIndex()) + ")";
							cell.setCellFormula(sumFormula);
							break;
						case 6 : 
							sumFormula = "SUM(G" + (cell.getRowIndex() - 3) + ":G" + (cell.getRowIndex()) + ")";
							cell.setCellFormula(sumFormula); 
							break;
						default : break;
						}
					}
				}
				
				if(skip && rt == 8) {
					row.getCell(0, Row.CREATE_NULL_AS_BLANK).setCellValue(fullName);
				}
			}
		}
		
//		<!-- Calculate Grade for TSR -->
		Cell sumScoreCell = sheet.getRow(sheet.getLastRowNum()).getCell(getColumnIndex("G"), Row.CREATE_NULL_AS_BLANK);
		Cell gradeCell = sheet.getRow(sheet.getLastRowNum() - 4).createCell(getColumnIndex("H"), Cell.CELL_TYPE_STRING);
		FormulaEvaluator evaluator = sheet.getWorkbook().getCreationHelper().createFormulaEvaluator();
		Double score = skip ? null : evaluator.evaluate(sumScoreCell).getNumberValue();
		gradeCell.setCellValue(getGrade(score));
		
//		<!-- Keep TSR Grade result for DSM Section -->
		if(tsrABGradeByDsmMap.get(dsmCode) == null) {
			tsrABGradeByDsmMap.put(dsmCode, new Double[]{0D, 0D});
		}
		Double[] gradeAB = tsrABGradeByDsmMap.get(dsmCode);
		gradeAB[0] += (getGrade(score).equals("A") || getGrade(score).equals("B") ? 1D : 0D);
		gradeAB[1] += 1;
	}
	

	private void sumRetention(Map<String, Double[]> retentionMap, String key, Double begin, Double during, Double end) {
		if(retentionMap.get(key) == null) retentionMap.put(key, new Double[]{0D, 0D, 0D});
		Double[] retentions = retentionMap.get(key);
		retentions[0] += begin;
		retentions[1] += during;
		retentions[2] += end;
	}

	private List<KpiCategorySetup> getKpiCategory(String tsrCode, String tsrLevel, String campaignCode, String listLotCode, Date effectiveDate, Date endDate) throws Exception {
		DetachedCriteria criteria = DetachedCriteria.forClass(KpiCategorySetup.class);
		
		if(tsrLevel.equals("TSR")) {
			criteria.add(Restrictions.isNull("tsr.tsrCode"));
			
			if(StringUtils.isBlank(listLotCode)) {
				criteria.add(Restrictions.isNull("listLotCode"));
			} else {
				criteria.add(Restrictions.eq("listLotCode", listLotCode));
			}

			criteria.add(Restrictions.eq("campaign.campaignCode", campaignCode));
		} else if(tsrLevel.equals("SUP")) {
			criteria.add(Restrictions.eq("tsr.tsrCode", tsrCode));
			
			if(StringUtils.isBlank(listLotCode) || StringUtils.isEmpty(listLotCode)) {
				criteria.add(Restrictions.isNull("listLotCode"));
			} else {
				criteria.add(Restrictions.eq("listLotCode", listLotCode));
			}

			criteria.add(Restrictions.eq("campaign.campaignCode", campaignCode));
		} else {
			criteria.add(Restrictions.eq("tsr.tsrCode", tsrCode));
			criteria.add(Restrictions.isNull("campaign.campaignCode"));
			criteria.add(Restrictions.isNull("listLotCode"));
			
		}
		criteria.add(Restrictions.eq("tsrLevel", tsrLevel));
//		criteria.add(Restrictions.eq("effectiveDate", effectiveDate));
//		criteria.add(Restrictions.eq("endDate", endDate));
		criteria.add(Restrictions.eq("effectiveDate", new java.sql.Date(effectiveDate.getTime())));
		criteria.add(Restrictions.eq("endDate", new java.sql.Date(endDate.getTime())));
		criteria.addOrder(Order.asc("id"));
		return kpiCategorySetupService.findByCriteria(criteria);
	}
	
	private Object[] getDsmKpi(String yearMonth, String dsmCode) throws Exception {
		/*
		 * <!-- Query Pattern -->
		 * select t.DSM_CODE
				, sum(t.TOTAL_AFYP) as TOTAL_AFYP
				, sum(t.FIRST_CONFIRM_SALE) as FIRST_CONFIRM_SALE
				, sum(t.ALL_SALE) as ALL_SALE
			from KPI_RESULT t
			where t.YEAR_MONTH = '201504'
			and t.DSM_CODE = '603353'
			group by t.DSM_CODE
		 */
		DetachedCriteria criteria = DetachedCriteria.forClass(KpiResult.class);
		criteria.add(Restrictions.eq("yearMonth", yearMonth));
		criteria.add(Restrictions.eq("dsm.tsrCode", dsmCode));
		
		criteria.setProjection(Projections.projectionList()
				.add(Projections.groupProperty("dsm.tsrCode"))
				.add(Projections.sum("totalAfyp"))
				.add(Projections.sum("firstConfirmSale"))
				.add(Projections.sum("allSale")));
		
		criteria.addOrder(Order.asc("dsm.tsrCode"));
		
		List<?> list = kpiResultService.findByCriteria(criteria);
		if(list.isEmpty()) {
			logger.error("Error KPI data for DSM is not found: Year Month: " + yearMonth + " | DSM Code: " + dsmCode);
			return null;
		}
		return (Object[]) list.get(0);
	}

	private Object[] getSupKpiByCampaign(String yearMonth, String campaignCode, String dsmCode, String supCode) throws Exception {
		/*
		 * <!-- Query Pattern -->
		 * select t.CAMPAIGN_CODE, t.LIST_LOT_CODE, t.DSM_CODE, t.SUP_CODE
		 		, sum(t.TOTAL_AFYP) as TOTAL_AFYP
				, count(t.TSR_CODE) as COUNT_TSR
				, sum(t.FIRST_CONFIRM_SALE) as FIRST_CONFIRM_SALE
				, sum(t.ALL_SALE) as ALL_SALE
				, sum(t.SUCCESS_POLICY) as SUCCESS_POLICY
				, sum(t.TOTAL_USED) as TOTAL_USED
			from KPI_RESULT t
			where 1 = 1
			and t.YEAR_MONTH = '201504'
			and t.CAMPAIGN_CODE = '141PA1715S03'
			and (t.LIST_LOT_CODE = 'AGD15' or t.LIST_LOT_CODE is null)
			and t.DSM_CODE = '603238'
			and t.SUP_CODE = '603391'
			group by t.CAMPAIGN_CODE, t.LIST_LOT_CODE, t.DSM_CODE, t.SUP_CODE
			order by t.CAMPAIGN_CODE, t.LIST_LOT_CODE, t.DSM_CODE, t.SUP_CODE
		 */
		
		String c = "", l = "";
		if(campaignCode.contains("_")) {
			c = campaignCode.substring(0, campaignCode.indexOf("_"));
			l = campaignCode.substring(campaignCode.indexOf("_") + 1, campaignCode.length());
		} else {
			c = campaignCode;
		}
		
		DetachedCriteria criteria = DetachedCriteria.forClass(KpiResult.class);
		criteria.add(Restrictions.eq("yearMonth", yearMonth));
		criteria.add(Restrictions.eq("campaign.campaignCode", c));
		criteria.add(Restrictions.disjunction().add(Restrictions.isNull("listLotCode")).add(Restrictions.eq("listLotCode", l)));
		criteria.add(Restrictions.eq("dsm.tsrCode", dsmCode));
		criteria.add(Restrictions.eq("tsm.tsrCode", supCode));
		
		criteria.setProjection(Projections.projectionList()
				.add(Projections.groupProperty("campaign.campaignCode"), "campaignCode")
				.add(Projections.groupProperty("listLotCode"), "listLotCode")
				.add(Projections.groupProperty("dsm.tsrCode"), "dsmCode")
				.add(Projections.groupProperty("tsm.tsrCode"), "tsmCode")
				.add(Projections.sum("totalAfyp"), "totalAfyp")
				.add(Projections.count("tsr.tsrCode"), "countTsr")
				.add(Projections.sum("firstConfirmSale"), "firstConfirmSale")
				.add(Projections.sum("allSale"), "allSale")
				.add(Projections.sum("successPolicy"), "successPolicy")
				.add(Projections.sum("totalUsed"), "totalUsed"));
		
		criteria.addOrder(Order.asc("campaign.campaignCode"));
		criteria.addOrder(Order.asc("listLotCode"));
		criteria.addOrder(Order.asc("dsm.tsrCode"));
		criteria.addOrder(Order.asc("tsm.tsrCode"));
		
		List<?> list = kpiResultService.findByCriteria(criteria);
		if(list.isEmpty()) {
			logger.error("Error Kpi data for SUP not found => yearMonth: " + yearMonth + " | campaignCode: " + c + " | listLotCode: " + l + " | dsmCode: " + dsmCode + " | supCode: " + supCode);
			return null;
		}
		return (Object[]) list.get(0);
	}
	
	private Object[] getTsrKpiByCampaign(String yearMonth, String campaignCode, String dsmCode, String supCode, String tsrCode) throws Exception {
		/*
		 * <!-- Query Pattern -->
		 * select t.CAMPAIGN_CODE, t.LIST_LOT_CODE, t.DSM_CODE, t.SUP_CODE, t.TSR_CODE
				, sum(t.TOTAL_AFYP) as TOTAL_AFYP
				, sum(t.ATTENDANCE) as ATTENDANCE
				, sum(t.TOTAL_TALK_HRS) as TOTAL_TALK_HRS
				, sum(t.FIRST_CONFIRM_SALE) as FIRST_CONFIRM_SALE
				, sum(t.ALL_SALE) as ALL_SALE
			from KPI_RESULT t
			where 1 = 1
			and t.YEAR_MONTH = '201504'
			and t.CAMPAIGN_CODE = '021DP1715L01'
			and (t.LIST_LOT_CODE = '' or t.LIST_LOT_CODE is null)
			and t.DSM_CODE = '602046'
			and t.SUP_CODE = '602517'
			and t.TSR_CODE = '602094'
			group by t.CAMPAIGN_CODE, t.LIST_LOT_CODE, t.DSM_CODE, t.SUP_CODE, t.TSR_CODE
			order by t.CAMPAIGN_CODE, t.LIST_LOT_CODE, t.DSM_CODE, t.SUP_CODE, t.TSR_CODE
		 */
		
		String c = "", l = "";
		if(campaignCode.contains("_")) {
			c = campaignCode.substring(0, campaignCode.indexOf("_"));
			l = campaignCode.substring(campaignCode.indexOf("_") + 1, campaignCode.length());
		} else {
			c = campaignCode;
		}
		
		DetachedCriteria criteria = DetachedCriteria.forClass(KpiResult.class);
		criteria.add(Restrictions.eq("yearMonth", yearMonth));
		criteria.add(Restrictions.eq("campaign.campaignCode", c));
		criteria.add(Restrictions.disjunction().add(Restrictions.isNull("listLotCode")).add(Restrictions.eq("listLotCode", l)));
		criteria.add(Restrictions.eq("dsm.tsrCode", dsmCode));
		criteria.add(Restrictions.eq("tsm.tsrCode", supCode));
		criteria.add(Restrictions.eq("tsr.tsrCode", tsrCode));
		
		criteria.setProjection(Projections.projectionList()
				.add(Projections.groupProperty("campaign.campaignCode"), "campaignCode")
				.add(Projections.groupProperty("listLotCode"), "listLotCode")
				.add(Projections.groupProperty("dsm.tsrCode"), "dsmCode")
				.add(Projections.groupProperty("tsm.tsrCode"), "tsmCode")
				.add(Projections.groupProperty("tsr.tsrCode"), "tsrCode")
				.add(Projections.sum("totalAfyp"), "totalAfyp")
				.add(Projections.sum("attendance"), "attendance")
				.add(Projections.sum("totalTalkHrs"), "totalTalkHrs")
				.add(Projections.sum("firstConfirmSale"), "firstConfirmSale")
				.add(Projections.sum("allSale"), "allSale"));
		
		criteria.addOrder(Order.asc("campaign.campaignCode"));
		criteria.addOrder(Order.asc("listLotCode"));
		criteria.addOrder(Order.asc("dsm.tsrCode"));
		criteria.addOrder(Order.asc("tsm.tsrCode"));
		criteria.addOrder(Order.asc("tsr.tsrCode"));
		
		List<?> list = kpiResultService.findByCriteria(criteria);
		if(list.isEmpty()) {
			logger.error("Error Kpi data for TSR not found => yearMonth: " + yearMonth + " | campaignCode: " + c + " | listLotCode: " + l + " | dsmCode: " + dsmCode + " | supCode: " + supCode + " | tsrCode: " + tsrCode);
			return null;
		}
		return (Object[]) list.get(0);
	}
	
	private List<KpiResult> getKpiResultByYearMonth(String yyyyMM) throws Exception {
		DetachedCriteria criteria = DetachedCriteria.forClass(KpiResult.class);
		criteria.add(Restrictions.eq("yearMonth", yyyyMM));
		criteria.addOrder(Order.asc("campaign.campaignCode"))
			.addOrder(Order.asc("listLotCode"))
			.addOrder(Order.asc("dsm.tsrCode"))
			.addOrder(Order.asc("tsm.tsrCode"))
			.addOrder(Order.asc("tsr.tsrCode"));
		return getKpiResultByCriteria(criteria);
	}
	
	private List<KpiResult> getKpiResultByCriteria(DetachedCriteria criteria) throws Exception {
		return kpiResultService.findByCriteria(criteria);
	}
	
	private List<KpiRetention> getKpiRetentions(String yyyyMM) throws Exception {
		KpiRetentionService service = (KpiRetentionService) AppConfig.getInstance().getBean("kpiRetentionService");
		return service.findByNamedQuery("execKpiRetentionByMonth", yyyyMM);
	}
	
	private InputStream getTemplateStream() {
		return Thread.currentThread().getContextClassLoader().getResourceAsStream(templateKpiReportPath);
	}

	private String getGrade(Double score) {
		String grade = "N/A";
		if(score != null) {
			if(score > 0.9D) {
				grade = "A";
			} else if(score > 0.8D) {
				grade = "B";
			} else if(score > 0.6D) {
				grade = "C";
			} else if(score <= 0.6D) {
				grade = "D";
			}
		}
		return grade;
	}
	
	private Integer getWorkDays(String yyyyMM) throws Exception {
		Integer days = 365;
		Date begin = DateUtil.convStringToDate("yyyyMMdd", yyyyMM + "01");
		Date end = null;
		
		Calendar cal = DateUtil.getCurrentCalendar();
		cal.setTime(begin);
		DateUtil.addMonth(cal, 1);
		DateUtil.addDay(cal, -1);
		end = DateUtil.convStringToDate("yyyyMMdd", DateUtil.convDateToString("yyyyMMdd", cal.getTime()));
		
		DetachedCriteria criteria = DetachedCriteria.forClass(Sales.class);
		criteria.add(Restrictions.between("saleDate", begin, end));
		criteria.setProjection(Projections.distinct(Projections.property("saleDate")));
		
		SalesService service = (SalesService) AppConfig.getInstance().getBean("salesService");
		List<?> list = service.findByCriteria(criteria);
		Object[] objs = list.toArray();
		for(Object obj : objs) {
			java.sql.Date date = (java.sql.Date) obj;
			if(isSaturdayOrSunDay(new Date(date.getTime()))) {
				list.remove(obj);
			}
		}
		
		days = list.size();
		return days;
	}
	
	private int getColumnIndex(String columnString) {
		return CellReference.convertColStringToIndex(columnString);
	}

	
	
	private boolean isMsigWB(String listLotDelim, String listLotCode) {
		return StringUtils.isNoneBlank(listLotCode) && listLotDelim.contains(listLotCode) ? true : false;
	}

	private boolean isSaturdayOrSunDay(Date date) {
		Calendar cal = DateUtil.getCurrentCalendar();
		cal.setTime(date);
		int day = cal.get(Calendar.DAY_OF_WEEK);
		if(day == Calendar.SATURDAY || day == Calendar.SUNDAY) {
			return true;
		}
		return false;
	}
	
	private void writeout(Workbook wb, String outPath, String fileName, int numOfTemplates) {
			OutputStream os = null;
			
			try {
				FileUtil.getInstance().createDirectory(outPath);
				String fullPath = outPath + File.separatorChar + fileName;
				
	//			<!-- remove template -->
				for(int i = 0; i < numOfTemplates; i++) wb.removeSheetAt(0);
				
				os = new FileOutputStream(fullPath);
				wb.write(os);
				logger.info("Writed to: " + fullPath);
			} catch(Exception e) {
				logger.error(e.getMessage(), e);
			} finally {
				try {os.close();} catch (IOException e) {}
				try {wb.close(); wb = null;} catch (IOException e) {}
			}
		}

	private void copyCellValue(Cell origCell, Cell toCell) {
		switch(origCell.getCellType()) {
		case Cell.CELL_TYPE_BLANK :
			toCell.setCellValue(origCell.getStringCellValue());
			break;
		case Cell.CELL_TYPE_BOOLEAN :
			toCell.setCellValue(origCell.getBooleanCellValue());
			break;
		case Cell.CELL_TYPE_ERROR :
			toCell.setCellValue(origCell.getErrorCellValue());
			break;
		case Cell.CELL_TYPE_FORMULA :
			toCell.setCellValue(origCell.getCellFormula());
			break;
		case Cell.CELL_TYPE_NUMERIC :
			toCell.setCellValue(origCell.getNumericCellValue());
			break;
		case Cell.CELL_TYPE_STRING :
			toCell.setCellValue(origCell.getRichStringCellValue());
			break;
		}
	}
	
	private void copyRowCellDataWithSameColumn(Sheet origSheet, Sheet toSheet, int origRowNum, int toRowNum, int startCellNum, int endCellNum, boolean isCopyValue) {
		Row toRow = toSheet.getRow(toRowNum) == null ? toSheet.createRow(toRowNum) : toSheet.getRow(toRowNum);
		Row origRow = origSheet.getRow(origRowNum);
		
		for(int i = startCellNum; i <= endCellNum; i++) {
			try {
				Cell origCell = origRow.getCell(i, Row.CREATE_NULL_AS_BLANK);
				Cell toCell = toRow.createCell(i, origCell.getCellType());
				if(isCopyValue) copyCellValue(origCell, toCell);
				toCell.setCellStyle(origCell.getCellStyle());
			} catch(Exception e) {
				logger.error("i: " + i + " | origRowNum: " + origRowNum);
				throw e;
			}
		}
	}
}
