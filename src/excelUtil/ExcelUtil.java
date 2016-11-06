package excelUtil;
/** 
 * @author HPC
 * 
 */
import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.List;

import org.apache.log4j.Logger;
import org.apache.poi.ss.usermodel.*;


public class ExcelUtil {

	//%%%%%%%%-------字段部分 开始----------%%%%%%%%%
	/**
	 * Excel文件路径
	 */
	private String excelPath = "";
	
	/**
	 * Excel文件
	 */
	private File excelFile;
	
	/**
	 *文件后缀名(xls,xlsx) 
	 */
	private String extension = "";
	
	/**
	 * 操作Excel的Workbook工具
	 */
	private Workbook workbook = null;
	
	/**
	 * 日志输出对象
	 */
	private static Logger logger = Logger.getLogger("excelLog");
	
	/**
	 * 最近一次读取的sheet列表
	 */
	public ArrayList<SheetReadUtil> sheetList = null; 

	
	/**
	 * 默认Excel内容的开始比较列位置为第一列（索引值为0）
	 */
	private int compareColumnPos = 0;

	/**
	 * 多文件合并时遇到名称重复时是否进行覆盖，
	 * 默认为true
	 */
	private boolean isOverWrite = true;
	
	/**
	 * 多文件合并时是否需要做内容比较(即相同的内容不重复出现),默认值为true
	 * (仅当不覆写目标内容是有效，即isOverWrite=false时有效)
	 */
	private boolean isNeedCompare = true;
	

	
	
	//%%%%%%%%-------字段部分 结束----------%%%%%%%%%
	

	/**
	 * @param excelPath  文件路径
	 * @throws IllegalArgumentException 文件不存在或格式错误
	 * @throws NullPointerException 文件路径为null
	 * @throws SecurityException 文件拒绝访问
	 */
	public ExcelUtil(String excelPath) throws IllegalArgumentException,NullPointerException,
											SecurityException{
		try {
			if(isExcelFile(excelPath)){
				this.excelPath = excelPath;
				this.excelFile = new File(excelPath);
				String name = this.excelFile.getName();
				this.extension = name.substring(name.lastIndexOf("."));
				FileInputStream is = new FileInputStream(excelFile); 
				//这种方式 Excel 2003/2007/2010 都是可以处理的  
		        this.workbook = WorkbookFactory.create(is) ;
			}else {
				throw new IllegalArgumentException("文件不是Excel文件");
			}
		}catch (IllegalArgumentException e) {
			throw e;
		}catch (NullPointerException e) {
			throw e;
		}catch (SecurityException e) {
			throw e;
		}catch (Exception e) {
			logger.error("other exception in ExcelUtil()", e);
		}
	}
	
	/**
	 * 重新new一个新的对象并返回
	 * @return
	 */
/*	public ExcelUtil returnNewInstance(){
		try {
			ExcelUtil instance = new  ExcelUtil(this.excelPath);
			return instance;
		} catch (Exception e) {
			logger.error("",e);
			return null;
		}
	}*/
	
	/**
	 * 检查文件是否为Excel文件
	 * @param filePath 文件路径
	 * @return 如果文件为Excel格式则返回true,否则false
	 * @throws IllegalArgumentException 文件不存在
	 * @throws NullPointerException 文件路径为null
	 * @throws SecurityException 文件拒绝访问
	 */
	public static boolean isExcelFile(String filePath) throws IllegalArgumentException,
															NullPointerException,SecurityException{
		try {
			File file = new File(filePath);
			if(!file.exists()){
				throw new IllegalArgumentException("路径错误,文件不存在");
			}else{
				String name = file.getName();
				String ext = name.substring(name.lastIndexOf("."));
				if(ext.equals(".xls") || ext.equals(".xlsx")) return true;
				else return false;
			}
		} catch (NullPointerException e) {
			throw new NullPointerException("文件路径为空");
		} catch (SecurityException  e) {	
			throw new SecurityException("文件拒绝访问");
		}
	}
	
	/**
	 * 读取第一个sheet
	 */
	public void readFirstSheet(){
		this.sheetList = changeSL2SBL(getSheetList(0, 1));
		readAllRows(sheetList.get(0));
	}
	
	/**
	 * 读取下标为Index的Sheet
	 * @param index sheet的下标值
	 * @throws IndexOutOfBoundsException 参数越界错误
	 */
	public void readSheetByIndex(int index) throws IndexOutOfBoundsException{
		try {
			this.sheetList = changeSL2SBL(getSheetList(index, 1));
			readAllRows(sheetList.get(0));
		} catch (IndexOutOfBoundsException e) {
			throw e;
		}
	}
	
	/**
	 * 通过名称读取sheet
	 * @param name sheet的名称
	 * @throws IllegalArgumentException 名称错误，无法获取指定sheet
	 */
	public void readSheetByName(String name) throws IllegalArgumentException{
		Sheet sheet = workbook.getSheet(name);
		if(sheet != null){
			sheetList.add(new SheetReadUtil(sheet));
			readAllRows(sheetList.get(0));
		}
		else
			throw new IllegalArgumentException("无法获取指定名称的sheet");
	}
	
	/**
	 * 读取指定范围的sheet列表
	 * @param startIndex sheet开始的下标值
	 * @param length 要读取sheets的长度
	 * @throws IndexOutOfBoundsException 参数越界错误
	 */
	public void readSheetList(int startIndex,int length) throws IndexOutOfBoundsException{
		try {
			this.sheetList = changeSL2SBL(getSheetList(startIndex, length));
			for(SheetReadUtil bean : sheetList){
				readAllRows(bean);
			}
		} catch (IndexOutOfBoundsException e) {
			throw e;
		}
	}
	
	/**
	 * 读取所有的sheet
	 */
	public void readAllSheet(){
		int sheetCount = workbook.getNumberOfSheets();
        SheetReadUtil sheetBean = null;
        for(int i=0;i<sheetCount;i++){
        	sheetBean = new SheetReadUtil(workbook.getSheetAt(i));
        	readAllRows(sheetBean);
        	sheetList.add(sheetBean);
        }
	}
	
	/**
	 * 获取需要的sheet列表
	 * @param startIndex sheet开始的下标值
	 * @param length 要读取sheets的长度,如果长度过长则读取至数据结尾
	 * @return Sheet 列表
	 * @throws IndexOutOfBoundsException 参数错误
	 */
	public ArrayList<Sheet> getSheetList(int startIndex,int length) throws IndexOutOfBoundsException{
		ArrayList<Sheet> sheetList = new ArrayList<>();
        int sheetCount = workbook.getNumberOfSheets();  //Sheet的数量  
        try {
			int endIndex = isIndexOutOfBounds(sheetCount, startIndex, length);
			for(int i=startIndex;i<=endIndex;i++){
	        	sheetList.add(workbook.getSheetAt(i));
	        }
	        return sheetList;
		} catch (IndexOutOfBoundsException e) {
			throw e;
		}
	}

	/**
	 * 将SheetList转换为SheetBeanList
	 * @param sheetList
	 * @return sheetBean列表
	 */
	private ArrayList<SheetReadUtil> changeSL2SBL(ArrayList<Sheet> sheetList){
		ArrayList<SheetReadUtil> sheetBeans = new ArrayList<>();
		SheetReadUtil bean = null;
		for(Sheet s:sheetList){
			bean = new SheetReadUtil(s);
			sheetBeans.add(bean);
		}
		return sheetBeans;
	}
	
	/**
	 * 读取SheetBean中所有的行
	 * @param sheetBean 要读取的sheetBean
	 */
	private void readAllRows(SheetReadUtil sheetBean){
		Sheet sheet = sheetBean.getSheet();
		int rowsCount = sheet.getLastRowNum() + 1;
		readRows(sheetBean, 0, rowsCount);
	}

	/**
	 * 读取sheetBean中指定的行
	 * @param sheetBean 要读取的sheetBean
	 * @param startIndex 开始的行下标
	 * @param length 读取长度
	 * @throws IndexOutOfBoundsException 参数越界错误
	 */
	private void readRows(SheetReadUtil sheetBean,int startIndex,int length) throws IndexOutOfBoundsException{
		Sheet sheet = sheetBean.getSheet();
		int rowsCount = sheet.getLastRowNum() + 1;
		try {
			int endIndex = isIndexOutOfBounds(rowsCount, startIndex, length);
			for(int i = startIndex;i<=endIndex;i++){
				if(endIndex == 0){//只读sheet的第一行
					Row r = sheet.getRow(0);
					if(r != null)
						sheetBean.addRow(r);
				}else{
					sheetBean.addRow(sheet.getRow(i));
				}
			}
		} catch (IndexOutOfBoundsException e) {
			throw e;
		}
	}
	
	/**
	 * 判断数据长度、起始下标和读取长度参数是否越界
	 * @param count 数据总长度,不能小于1
	 * @param startIndex 起始下标不能小于零或大于最大值
	 * @param length 读取的长度,不能小于0
	 * @return 如果参数越界抛出异常,否则返回要读取的最后一个下标值(如果下标越界,则返回最大下标值)
	 * @throws IndexOutOfBoundsException 
	 */
	public static int isIndexOutOfBounds(int count,int startIndex,int length) throws IndexOutOfBoundsException{
		if(count<1){
			throw new IndexOutOfBoundsException("数据长度小于1");
		}
		if(length<0){
        	throw new IndexOutOfBoundsException("读取长度小于零");
        }
        if(startIndex > count -1 || startIndex < 0){
        	throw new IndexOutOfBoundsException("开始下标大于最大的下标值或小于零");
        }
        //要读取的最后一个下标,如果下标越界，则读取至最后一个值
        int endIndex = startIndex + length - 1;
		if (endIndex >= count)
			endIndex = count - 1;
        return endIndex;
	}
	
	





	

	
	/**
	 * 修改Excel，并另存为
	 * 
	 * @Title: WriteExcel
	 * @Date : 2014-9-11 下午01:33:59
	 * @param wb
	 * @param rowList
	 * @param xlsPath
	 */
	/*private void writeExcel(Workbook wb, List<Row> rowList, String xlsPath) {

		
		Sheet sheet = wb.getSheetAt(0);// 修改第一个sheet中的值

		// 如果每次重写，那么则从开始读取的位置写，否则果获取源文件最新的行。
		int lastRowNum = isOverWrite ? startReadRowPos : sheet.getLastRowNum() + 1;
		int t = 0;//记录最新添加的行数
		//out("要添加的数据总条数为："+rowList.size());
		for (Row row : rowList) {
			if (row == null) continue;
			// 判断是否已经存在该数据
			int pos = findInExcel(sheet, row);

			Row r = null;// 如果数据行已经存在，则获取后重写，否则自动创建新行。
			if (pos >= 0) {
				sheet.removeRow(sheet.getRow(pos));
				r = sheet.createRow(pos);
			} else {
				r = sheet.createRow(lastRowNum + t++);
			}
			
			//用于设定单元格样式
			CellStyle newstyle = wb.createCellStyle();
			
			//循环为新行创建单元格
			for (int i = row.getFirstCellNum(); i < row.getLastCellNum(); i++) {
				Cell cell = r.createCell(i);// 获取数据类型
				cell.setCellValue(getCellValue(row.getCell(i)));// 复制单元格的值到新的单元格
				// cell.setCellStyle(row.getCell(i).getCellStyle());//出错
				if (row.getCell(i) == null) continue;
				copyCellStyle(row.getCell(i).getCellStyle(), newstyle); // 获取原来的单元格样式
				cell.setCellStyle(newstyle);// 设置样式
				// sheet.autoSizeColumn(i);//自动跳转列宽度
			}
		}
		//out("其中检测到重复条数为:" + (rowList.size() - t) + " ，追加条数为："+t);
		
		// 统一设定合并单元格
		setMergedRegion(sheet);
		
		try {
			// 重新将数据写入Excel中
			FileOutputStream outputStream = new FileOutputStream(xlsPath);
			wb.write(outputStream);
			outputStream.flush();
			outputStream.close();
		} catch (Exception e) {
			//out("写入Excel时发生错误！ ");
			e.printStackTrace();
		}
	}*/



}