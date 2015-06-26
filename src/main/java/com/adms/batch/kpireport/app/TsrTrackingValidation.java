package com.adms.batch.kpireport.app;

import java.io.File;
import java.io.FileInputStream;
import java.io.FilenameFilter;
import java.io.InputStream;
import java.util.ArrayList;
import java.util.List;

import org.apache.commons.lang3.StringUtils;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.ss.util.CellReference;

import com.adms.support.FileWalker;
import com.adms.utils.FileUtil;
import com.adms.utils.Logger;


public class TsrTrackingValidation {

	private static Logger logger = Logger.getLogger();
	
	private static final String OK = "OK";
	private static final String ERR = "ERR";
	
	public static void main(String[] args) {
		List<String[]> records = new ArrayList<>();
		String rootPath = args[0];
		
		String outDir = args[1];
		String outName = args[2];
		
		InputStream wbIS = null;
		try {
			logger.setLogFileName(args[3]);

			logger.info("Start");
			List<String> dirs = getFilePaths(rootPath);
			
			for(String dir : dirs) {
				String listLotCode = null;
				String valid = null;
				
				wbIS = new FileInputStream(dir);
				
				Workbook wb = WorkbookFactory.create(wbIS);
				boolean isSingleSheet = wb.getNumberOfSheets() == 1;
				
				for(int sheetIdx = 0; sheetIdx < wb.getNumberOfSheets(); sheetIdx++) {
					
					Sheet sheet = wb.getSheetAt(sheetIdx);
					String listLotName = sheet.getRow(1).getCell(CellReference.convertColStringToIndex("C"), Row.CREATE_NULL_AS_BLANK).getStringCellValue();
					
					if(listLotName.contains(",")) {
						valid = isSingleSheet ? ERR : sheetIdx == 0 ? OK : ERR;
					} else {
						listLotCode = getListLotCode(listLotName);
//						<!-- check -->
						valid = StringUtils.isBlank(listLotCode) ? ERR : OK;
					}

					if(valid.equals(ERR)) 
						records.add(new String[]{valid, dir, wb.getSheetName(sheetIdx), listLotName});
				}
			}
			
			writeDataOut(outDir, outName, records);
			
		} catch(Exception e) {
			e.printStackTrace();
		} finally {
			try { wbIS.close(); } catch(Exception e) {}
		}
		logger.info("Finish");
	}
	
	private static void writeDataOut(String outDir, String outName, List<String[]> records) throws Exception {
		FileUtil.getInstance().createDirectory(outDir);
		
		StringBuffer buffer = new StringBuffer();
		
		int row = 0;
		int idx = 0;
		
		for(String[] strs : records) {
			idx = 0;
			
			for(String str : strs) {
				buffer.append(toCoveredTxt(str));
				if(idx < strs.length) {
					buffer.append(",");
				}
			}
			
			if(row < records.size()) {
				buffer.append("\n");
			}
			
			row++;
		}
		
		String msg = FileUtil.getInstance().writeout(new File(outDir + "/" + outName), buffer);
		logger.info("write to: " + msg);
	}
	
	private static String toCoveredTxt(String arg) {
		return new String("\"" + arg + "\"");
	}
	
	private static String getListLotCode(String val) {
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
				logger.error("Cannot Retrieve: " + val);
			}
		}
		return result;
	}
	
	private static List<String> getFilePaths(String root) {
		FileWalker fw = new FileWalker();
		
		fw.walk(root, new FilenameFilter() {

			@Override
			public boolean accept(File file, String name) {
				if(!name.contains("~$") 
						&& (name.contains("TsrTracking")
							|| name.contains("TSRTracking")
							|| name.contains("TSRTRA"))
						&& (!name.contains("CTD")
								&& !name.contains("MTD")
								&& !name.contains("_ALL")
								&& !name.contains("QA_Report")
								&& !name.contains("QC_Reconfirm")
								&& !name.contains("SalesReportByRecords")))
					return true;
				return false;
			}
		});
		
		return fw.getFileList();
	}
}

