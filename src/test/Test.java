package test;



import java.util.ArrayList;
import java.util.HashMap;
import java.util.Map;












import org.apache.log4j.Logger;
import org.apache.log4j.PropertyConfigurator;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.Row;






import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.util.CellRangeAddress;

import excelUtil.ExcelReadUtil;
import excelUtil.ExcelWriteUtil;
import excelUtil.SheetReadUtil;
import excelUtil.SheetWriteUtil;

public class Test {
	private static Logger logger = Logger.getLogger(Test.class);
	static ExcelReadUtil excelReadUtil = null;
	static ExcelWriteUtil excelWriteUtil = null;
	public static void main(String[] args) {
		PropertyConfigurator.configure("config/log4j.properties");
		
		
		try {
			String url1 = "C:/Users/Administrator/Desktop/测试原始数据表.xls";
			excelReadUtil = new ExcelReadUtil(url1);
			
			excelWriteUtil = new ExcelWriteUtil(
					"C:/Users/Administrator/Desktop", "测试生成表");

			//printAll(0);
			writeExcel();
			
			
			System.out.println("ALL DONE");
		} catch (Exception e) {
			// 
			logger.error("main()",e);
		}
		
		
		
	}
	
	public static void testTitle(){
		excelReadUtil.readSheetByIndex(1);
		SheetReadUtil bean = new SheetReadUtil(excelReadUtil.getSheetList().get(0));
		bean.readAllRows();
		
		System.out.println("Titles");
		//bean.setTitles(bean.getRowAt(2), 0, 10);
		HashMap<Integer, String> title = bean.getTitles();
		for(Map.Entry<Integer, String> entry:title.entrySet()){
			System.out.println("index:"+entry.getKey()+"  value:"+entry.getValue());
		}
		
		System.out.println("四的Index："+bean.getTitleColIndexByValue("四"));
	}

	
	public static void filter(){
		excelReadUtil.readSheetByIndex(0);
		SheetReadUtil bean = new SheetReadUtil(excelReadUtil.getSheetList().get(0));
		bean.readAllRows();
		System.out.println("row size: "+bean.getAllRowList().size());

		bean.setTitles(bean.getRowAt(1), 0, 38);
		
		ArrayList<Row> rows = bean.getRowListByArg(bean.getAllRowList(), "户口类型", "城镇户口");
		ArrayList<Row> rows2 = bean.getRowListByArg(rows, "政治面貌", "中共党员");
		
		String str = "";
		for(int i=0;i<rows2.size();i++){
			Row row = rows2.get(i);
			str = "";
			if(row == null){
				System.out.println("-------row is null-------------");
				continue;
			}
			for(Cell cell:bean.getOneRowAllCells(row)){
				if(cell == null){
					str += " [null] ";
					continue;
				}
				str += " ["+bean.getCellValue(cell) + "] ";
			}
			System.out.println(str);
		}
	}
	
	public static void printOneCol(int sheetIndex,int colIndex){
		excelReadUtil.readSheetByIndex(sheetIndex);
		SheetReadUtil bean = new SheetReadUtil(excelReadUtil.getSheetList().get(0));
		bean.readAllRows();
		
		ArrayList<Cell> cells = bean.getOneColumnAllCells(colIndex);
		String str = "";
		for(Cell c : cells){
			str = "";
			if(c == null){
				str = "[empty]";
				System.out.println(str);
				continue;
			}
			str = "["+bean.getCellValue(c)+"]";
			System.out.println(str);
		}
	}
	
	public static void printAll(int sheetIndex){
		excelReadUtil.readSheetByIndex(sheetIndex);
		SheetReadUtil bean = new SheetReadUtil(excelReadUtil.getSheetList().get(0));
		bean.readAllRows();
		System.out.println("row size: "+bean.getAllRowList().size());
		String str = "";
		for(Row row:bean.getAllRowList()){
			
			str = "";
			if(row == null){
				System.out.println("--------row is null----------------------");
				continue;
			}
			for(Cell cell:bean.getOneRowAllCells(row)){
				if(cell == null){
					str += " [null] ";
					continue;
				}
				str += " ["+bean.getCellValue(cell) + "] ";
			}
			System.out.println(str);
		}
	}
	
	public static void writeExcel(){
		try {
			excelReadUtil.readSheetList(0, 2);
			for (Sheet sheet : excelReadUtil.getSheetList()) {
				SheetReadUtil sRead = new SheetReadUtil(sheet);
				sRead.readAllRows();
				
				writeOneSheet(sRead);
			}
			excelWriteUtil.writeToExcel();
				
		} catch (Exception e) {
			logger.error("", e);
		}
		
	}
	
	@SuppressWarnings("deprecation")
	public static void writeOneSheet(SheetReadUtil sRead) {
		try {
			
			SheetWriteUtil sWrite = new SheetWriteUtil(
					excelWriteUtil.createSheet());

			if (sRead.hasMerged()) {// 添加合并区域
				int mergedCount = sRead.getSheet().getNumMergedRegions();
				for (int i = 0; i < mergedCount; i++) {
					CellRangeAddress rangeAddress = sRead.getSheet()
							.getMergedRegion(i);
					sWrite.addMergedRegion(rangeAddress.getFirstRow(),
							rangeAddress.getLastRow(),
							rangeAddress.getFirstColumn(),
							rangeAddress.getLastColumn());
				}
			}

			for (Row row : sRead.getAllRowList()) { // 添加单元格数据
				if (row == null) {
					continue;
				} else {
					ArrayList<Cell> cellList = sRead.getOneRowCellList(row, 0,
							21);
					for (Cell cell : cellList) {
						if (cell == null) {
							continue;
						} else {
							Cell c = sWrite.getValidCell(cell.getRowIndex(),
									cell.getColumnIndex());
							if (c == null) {// 此单元格无效，无需写入
								continue;
							} else {
								CellType type = null;
								type = cell.getCellTypeEnum();
								CellStyle style = null;
								style = sWrite.getCommonCellStyle_alignLeft();
								Font font = sWrite.getCommonFont_content();
								style.setFont(font);
								// style.setWrapText(true);
								String value = sRead.getCellValue(cell);

								sWrite.setAutoSizeColumn(cell.getColumnIndex());

								c.setCellType(type);
								c.setCellStyle(style);
								c.setCellValue(value);
							}
						}//cell != null
					}//end for cell
				}//row != null
			}//end row for
		} catch (Exception e) {
			// TODO: handle exception

		}

	}
	
}
