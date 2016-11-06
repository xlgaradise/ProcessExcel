package excel;
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
	public ArrayList<SheetReadBean> sheetList = null; 

	
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
			sheetList.add(new SheetReadBean(sheet));
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
			for(SheetReadBean bean : sheetList){
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
        SheetReadBean sheetBean = null;
        for(int i=0;i<sheetCount;i++){
        	sheetBean = new SheetReadBean(workbook.getSheetAt(i));
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
	private ArrayList<SheetReadBean> changeSL2SBL(ArrayList<Sheet> sheetList){
		ArrayList<SheetReadBean> sheetBeans = new ArrayList<>();
		SheetReadBean bean = null;
		for(Sheet s:sheetList){
			bean = new SheetReadBean(s);
			sheetBeans.add(bean);
		}
		return sheetBeans;
	}
	
	/**
	 * 读取SheetBean中所有的行
	 * @param sheetBean 要读取的sheetBean
	 */
	private void readAllRows(SheetReadBean sheetBean){
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
	private void readRows(SheetReadBean sheetBean,int startIndex,int length) throws IndexOutOfBoundsException{
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
	 * 自动根据文件扩展名，调用对应的写入方法
	 * 
	 * @Title: writeExcel
	 * @Date : 2014-9-11 下午01:50:38
	 * @param rowList
	 * @param xlsPath
	 * @throws IOException
	 */
	public void writeExcel(List<Row> rowList, String xlsPath) throws IOException {

		//扩展名为空时，
		if (xlsPath.equals("")){
			throw new IOException("文件路径不能为空！");
		}
		
		//获取扩展名
		//String ext = xlsPath.substring(xlsPath.lastIndexOf(".")+1);
		
		/*try {
			
			if("xls".equals(ext)){				//使用xls方式写入
				//writeExcel_xls(rowList,xlsPath);
			}else if("xls".equals(ext)){		//使用xlsx方式写入
				//writeExcel_xlsx(rowList,xlsPath);
			}else{									//依次尝试xls、xlsx方式写入
				//out("您要操作的文件没有扩展名，正在尝试以xls方式写入...");
				try{
					//writeExcel_xls(rowList,xlsPath);
				} catch (IOException e1) {
					//out("尝试以xls方式写入，结果失败！，正在尝试以xlsx方式读取...");
					try{
						//writeExcel_xlsx(rowList,xlsPath);
					} catch (IOException e2) {
						//out("尝试以xls方式写入，结果失败！\n请您确保您的文件是Excel文件，并且无损，然后再试。");
						throw e2;
					}
				}
			}
		} catch (IOException e) {
			throw e;
		}*/
	}


	/**
	 * 修改Excel（97-03版，xls格式）
	 * 
	 * @Title: writeExcel_xls
	 * @Date : 2014-9-11 下午01:50:38
	 * @param rowList
	 * @param src_xlsPath
	 * @param dist_xlsPath
	 * @throws IOException
	 */
	/*public void writeExcel_xls(List<Row> rowList, String src_xlsPath, String dist_xlsPath) throws IOException {

		// 判断文件路径是否为空
		if (dist_xlsPath == null || dist_xlsPath.equals("")) {
			//out("文件路径不能为空");
			throw new IOException("文件路径不能为空");
		}
		// 判断文件路径是否为空
		if (src_xlsPath == null || src_xlsPath.equals("")) {
			//out("文件路径不能为空");
			throw new IOException("文件路径不能为空");
		}

		// 判断列表是否有数据，如果没有数据，则返回
		if (rowList == null || rowList.size() == 0) {
			//out("文档为空");
			return;
		}

		try {
			HSSFWorkbook wb = null;

			// 判断文件是否存在
			File file = new File(dist_xlsPath);
			if (file.exists()) {
				// 如果复写，则删除后
				if (isOverWrite) {
					file.delete();
					// 如果文件不存在，则创建一个新的Excel
					// wb = new HSSFWorkbook();
					// wb.createSheet("Sheet1");
					wb = new HSSFWorkbook(new FileInputStream(src_xlsPath));
				} else {
					// 如果文件存在，则读取Excel
					wb = new HSSFWorkbook(new FileInputStream(file));
				}
			} else {
				// 如果文件不存在，则创建一个新的Excel
				// wb = new HSSFWorkbook();
				// wb.createSheet("Sheet1");
				wb = new HSSFWorkbook(new FileInputStream(src_xlsPath));
			}

			// 将rowlist的内容写到Excel中
			writeExcel(wb, rowList, dist_xlsPath);

		} catch (IOException e) {
			e.printStackTrace();
		}
	}
*/


	/**
	 * 修改Excel（2007版，xlsx格式）
	 * 
	 * @Title: writeExcel_xlsx
	 * @Date : 2014-9-11 下午01:50:38
	 * @param rowList
	 * @param xlsPath
	 * @throws IOException
	 */
	/*public void writeExcel_xlsx(List<Row> rowList, String src_xlsPath, String dist_xlsPath) throws IOException {

		// 判断文件路径是否为空
		if (dist_xlsPath == null || dist_xlsPath.equals("")) {
			//out("文件路径不能为空");
			throw new IOException("文件路径不能为空");
		}
		// 判断文件路径是否为空
		if (src_xlsPath == null || src_xlsPath.equals("")) {
			//out("文件路径不能为空");
			throw new IOException("文件路径不能为空");
		}

		// 判断列表是否有数据，如果没有数据，则返回
		if (rowList == null || rowList.size() == 0) {
			//out("文档为空");
			return;
		}

		try {
			// 读取文档
			XSSFWorkbook wb = null;

			// 判断文件是否存在
			File file = new File(dist_xlsPath);
			if (file.exists()) {
				// 如果复写，则删除后
				if (isOverWrite) {
					file.delete();
					// 如果文件不存在，则创建一个新的Excel
					// wb = new XSSFWorkbook();
					// wb.createSheet("Sheet1");
					wb = new XSSFWorkbook(new FileInputStream(src_xlsPath));
				} else {
					// 如果文件存在，则读取Excel
					wb = new XSSFWorkbook(new FileInputStream(file));
				}
			} else {
				// 如果文件不存在，则创建一个新的Excel
				// wb = new XSSFWorkbook();
				// wb.createSheet("Sheet1");
				wb = new XSSFWorkbook(new FileInputStream(src_xlsPath));
			}
			// 将rowlist的内容添加到Excel中
			writeExcel(wb, rowList, dist_xlsPath);

		} catch (IOException e) {
			e.printStackTrace();
		}
	}
*/


	

	
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

		if (wb == null) {
			//out("操作文档不能为空！");
			return;
		}

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

	/**
	 * 查找某行数据是否在Excel表中存在，返回行数。
	 * 
	 * @Title: findInExcel
	 * @Date : 2014-9-11 下午02:23:12
	 * @param sheet
	 * @param row
	 * @return
	 */
	/*private int findInExcel(Sheet sheet, Row row) {
		int pos = -1;

		try {
			// 如果覆写目标文件，或者不需要比较，则直接返回
			if (isOverWrite || !isNeedCompare) {
				return pos;
			}
			for (int i = startReadRowPos; i <= sheet.getLastRowNum() + endReadRowPos; i++) {
				Row r = sheet.getRow(i);
				if (r != null && row != null) {
					String v1 = getCellValue(r.getCell(compareColumnPos));
					String v2 = getCellValue(row.getCell(compareColumnPos));
					if (v1.equals(v2)) {
						pos = i;
						break;
					}
				}
			}
		} catch (Exception e) {
			e.printStackTrace();
		}
		return pos;
	}*/

	/**
	 * 复制一个单元格样式到目的单元格样式
	 * 
	 * @param fromStyle
	 * @param toStyle
	 */
	/*public static void copyCellStyle(CellStyle fromStyle, CellStyle toStyle) {
		toStyle.setAlignment(fromStyle.getAlignment());
		// 边框和边框颜色
		toStyle.setBorderBottom(fromStyle.getBorderBottom());
		toStyle.setBorderLeft(fromStyle.getBorderLeft());
		toStyle.setBorderRight(fromStyle.getBorderRight());
		toStyle.setBorderTop(fromStyle.getBorderTop());
		toStyle.setTopBorderColor(fromStyle.getTopBorderColor());
		toStyle.setBottomBorderColor(fromStyle.getBottomBorderColor());
		toStyle.setRightBorderColor(fromStyle.getRightBorderColor());
		toStyle.setLeftBorderColor(fromStyle.getLeftBorderColor());

		// 背景和前景
		toStyle.setFillBackgroundColor(fromStyle.getFillBackgroundColor());
		toStyle.setFillForegroundColor(fromStyle.getFillForegroundColor());

		// 数据格式
		toStyle.setDataFormat(fromStyle.getDataFormat());
		toStyle.setFillPattern(fromStyle.getFillPattern());
		// toStyle.setFont(fromStyle.getFont(null));
		toStyle.setHidden(fromStyle.getHidden());
		toStyle.setIndention(fromStyle.getIndention());// 首行缩进
		toStyle.setLocked(fromStyle.getLocked());
		toStyle.setRotation(fromStyle.getRotation());// 旋转
		toStyle.setVerticalAlignment(fromStyle.getVerticalAlignment());
		toStyle.setWrapText(fromStyle.getWrapText());

	}
*/
	/**
	 * 获取合并单元格的值
	 * 
	 * @param sheet
	 * @param row
	 * @param column
	 * @return
	 */
/*	public void setMergedRegion(Sheet sheet) {
		int sheetMergeCount = sheet.getNumMergedRegions();

		for (int i = 0; i < sheetMergeCount; i++) {
			// 获取合并单元格位置
			CellRangeAddress ca = sheet.getMergedRegion(i);
			int firstRow = ca.getFirstRow();
			if (startReadRowPos - 1 > firstRow) {// 如果第一个合并单元格格式在正式数据的上面，则跳过。
				continue;
			}
			int lastRow = ca.getLastRow();
			int mergeRows = lastRow - firstRow;// 合并的行数
			int firstColumn = ca.getFirstColumn();
			int lastColumn = ca.getLastColumn();
			// 根据合并的单元格位置和大小，调整所有的数据行格式，
			for (int j = lastRow + 1; j <= sheet.getLastRowNum(); j++) {
				// 设定合并单元格
				sheet.addMergedRegion(new CellRangeAddress(j, j + mergeRows, firstColumn, lastColumn));
				j = j + mergeRows;// 跳过已合并的行
			}

		}
	}
	*/
/*	private class SheetBean {
		private Sheet sheet;
		private ArrayList<Row> rows = null;

		public SheetBean(Sheet sheet){
			this.sheet = sheet;
		}

		public SheetBean(){
			
		}
	}*/

}