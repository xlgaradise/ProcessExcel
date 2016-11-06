package test;



import java.util.ArrayList;
import java.util.HashMap;
import java.util.Map;



import org.apache.log4j.Logger;
import org.apache.log4j.PropertyConfigurator;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;



import excel.ExcelUtil;
import excel.SheetReadBean;

public class Test {
	private static Logger logger = Logger.getLogger(Test.class);
	static ExcelUtil eu = null;
	public static void main(String[] args) {
		// TODO 自动生成的方法存根
		PropertyConfigurator.configure("config/log4j.properties");
		
		
		try {
			String url1 = "C:/Users/Administrator/Desktop/excel/软件学院2016-2017-1-研究生课程表.xls";
			String url2 = "C:/Users/Administrator/Desktop/excel/软件学院-16级研究生个人信息 (2).xls";
			eu = new ExcelUtil(url2);

			//dateFormat();
			//printAll();
			//filter();
			printCol();
			
			
			
			System.out.println("ALL DONE");
		} catch (Exception e) {
			// TODO 自动生成的 catch 块
			logger.error("main()",e);
		}
		
		
		
	}
	
	public static void dateFormat(){
		eu.readSheetByIndex(1);
		SheetReadBean bean = eu.sheetList.get(0);
		int startRowIndex = 0;
		int length = 1;
		
		String str = "";
		for(int i=startRowIndex;i<startRowIndex+length;i++){
			Row row = bean.getAllRowList().get(i);
			
			if(row == null){
				System.out.println("------------------------------");
				continue;
			}
			for(Cell cell:bean.getOneRowAllCells(row)){
				str = "";
				if(cell == null){
					str += "empty" + "---#";
					continue;
				}
				if(cell.getCellTypeEnum().toString().equals("NUMERIC")){
					str = "value:"+bean.getCellValue(cell) +
						"    format:"+cell.getCellStyle().getDataFormat();
					System.out.println(str);
				}
			}
			
		}
	}
	
	public static void title(){
		eu.readSheetByIndex(1);
		SheetReadBean bean = eu.sheetList.get(0);
		print(bean);
		
		System.out.println("Titles");
		//bean.setTitles(bean.getRowAt(2), 0, 10);
		HashMap<Integer, String> title = bean.getTitles();
		for(Map.Entry<Integer, String> entry:title.entrySet()){
			System.out.println("index:"+entry.getKey()+"  value:"+entry.getValue());
		}
		
		System.out.println("四的Index："+bean.getTitleColIndexByValue("四"));
	}
	
	public static void printAll(){
		eu.readSheetByIndex(0);
		SheetReadBean bean = eu.sheetList.get(0);
		System.out.println("row size: "+bean.getAllRowList().size());
		print(bean);
	}
	
	public static void filter(){
		eu.readSheetByIndex(0);
		SheetReadBean bean = eu.sheetList.get(0);
		System.out.println("row size: "+bean.getAllRowList().size());

		bean.setTitles(bean.getRowAt(1), 0, 38);
		
		ArrayList<Row> rows = bean.getRowListByArg(bean.getAllRowList(), "户口类型", "城镇户口");
		ArrayList<Row> rows2 = bean.getRowListByArg(rows, "政治面貌", "中共党员");
		
		String str = "";
		for(int i=0;i<rows2.size();i++){
			Row row = rows2.get(i);
			str = "";
			if(row == null){
				System.out.println("------------------------------");
				continue;
			}
			for(Cell cell:bean.getOneRowAllCells(row)){
				if(cell == null){
					str += "empty" + "---#";
					continue;
				}
				str += bean.getCellValue(cell) + "---#";
			}
			System.out.println(str);
		}
	}
	
	public static void printCol(){
		eu.readSheetByIndex(0);
		SheetReadBean bean = eu.sheetList.get(0);
		
		ArrayList<Cell> cells = bean.getOneColumnAllCells(1);
		String str = "";
		for(Cell c : cells){
			str = "";
			if(c == null){
				str = "empty";
				System.out.println(str);
				continue;
			}
			str = bean.getCellValue(c);
			System.out.println(str);
		}
	}
	
	public static void print(SheetReadBean bean){
		String str = "";
		for(Row row:bean.getAllRowList()){
			
			str = "";
			if(row == null){
				System.out.println("------------------------------");
				continue;
			}
			for(Cell cell:bean.getOneRowAllCells(row)){
				if(cell == null){
					str += "empty" + "---#";
					continue;
				}
				str += bean.getCellValue(cell) + "---#";
			}
			System.out.println(str);
		}
	}
	
	
	
}
