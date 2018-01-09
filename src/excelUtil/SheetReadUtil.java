package excelUtil;
/** 
 * @author HPC
 * 
 */
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Date;
import java.util.HashMap;
import java.util.Map;

import org.apache.poi.hssf.usermodel.HSSFDateUtil;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellValue;
import org.apache.poi.ss.usermodel.DateUtil;
import org.apache.poi.ss.usermodel.FormulaEvaluator;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.util.CellRangeAddress;

import excelUtil.CellTypeUtil.TypeEnum;
import exception.ExcelIllegalArgumentException;
import exception.ExcelIndexOutOfBoundsException;
import exception.ExcelNoTitleException;
import exception.ExcelNullParameterException;

/**
 *��sheet��Ĺ���,�ɻ�ȡ�С��С���Ԫ���
 */
public class SheetReadUtil {

	//%%%%%%%%-------�ֶβ��� ��ʼ----------%%%%%%%%%
	
	/**
	 * ʵ����sheet
	 */
	protected Sheet sheet;
	
	/**
	 * ��sheet��ָ��������
	 */
	protected ArrayList<Row> rowList = null;
	
	/**
	 * ��ѡ����(IntegerΪ��������±�,StringΪ��������)
	 */
	protected HashMap<Integer, String> titles;
	
	
	
	//%%%%%%%%-------�ֶβ��� ����----------%%%%%%%%%
	

	/**
	 * ������sheet�Ĺ��ߣ���ExcelReadUtil��ȡʵ����
	 * @param sheet ��Ҫ������sheetʵ��,��ExcelReadUtil��ȡ
	 * @throws ExcelNullParameterException ����Ϊnull
	 */
	public SheetReadUtil(Sheet sheet) throws ExcelNullParameterException{
		if(sheet == null){
			throw new ExcelNullParameterException();
		}
		this.sheet = sheet;
		rowList = new ArrayList<>();
		titles = new HashMap<>();
	}
	

	/**
	 * ��ȡʵ����sheet
	 * @return ���ظ�SheetReadUtil������sheet
	 */
	public Sheet getSheet(){
		return sheet;
	}
	
	/**
	 * ��ȡsheet����Ч������
	 * @return
	 */
	public int getPhysicalNumberOfRows(){
		return sheet.getPhysicalNumberOfRows();
	}
	
	/**
	 * ��ȡSheet�����е���
	 */
	public void readAllRows(){
		rowList.clear();
		int rowsCount = sheet.getLastRowNum() + 1;
		try {
			readRows(0, rowsCount);
		} catch (ExcelIndexOutOfBoundsException e) {
		}
	}

	/**
	 * ��ȡsheet��ָ������
	 * @param startIndex ��ʼ�����±�
	 * @param length ��ȡ����
	 * @throws ExcelIndexOutOfBoundsException ��ʼ�±겻��С�����������ֵ��length��ȡ�ĳ��Ȳ���С��0
	 */
	public void readRows(int startIndex,int length) throws ExcelIndexOutOfBoundsException{
		rowList.clear();
		int rowsCount = sheet.getLastRowNum() + 1;
		int endIndex = isIndexOutOfBounds(rowsCount, startIndex, length);
		for (int i = startIndex; i <= endIndex; i++) {
			if (endIndex == 0) {// ֻ��sheet�ĵ�һ��
				Row r = sheet.getRow(0);
				if (r != null)
					this.rowList.add(r);
			} else {
				this.rowList.add(sheet.getRow(i));
			}
		}
	}
	
	/**
	 * ����rowList�б�
	 * @return
	 */
	public ArrayList<Row> getRowList() {
		return rowList;
	}
	
	
	
	/**
	 * ͨ��ָ����������ֵ�����ݼ����ȡָ����RowList�б�
	 * @param rows Row�б����ݼ�
	 * @param title ��������
	 * @param value ��������ֵ
	 * @return ����ָ��RowList�б�
	 * @throws ExcelIllegalArgumentException ָ�����ⲻ����
	 * @throws ExcelNoTitleException δ���ñ���
	 */
	/*public ArrayList<Row> getRowListByArg(ArrayList<Row> rows,String title,String value) 
			throws ExcelIllegalArgumentException,ExcelNoTitleException{
		
		int columnIndex = getTitleColIndexByValue(title);
		ArrayList<Row> rowList = null;
		if(columnIndex != -1){//titleֵ����
			rowList = new ArrayList<>();
			for(Row row : rows){
				if(row == null){
					continue;
				}
				String v = getCellValue(row.getCell(columnIndex));
				if(v.equals(value)){
					rowList.add(row);
				}
			}
			return rowList;
		}else{
			throw new ExcelIllegalArgumentException();
		}
	}*/
	
	/**
	 * ͨ��ָ����������ֵ�����ݼ���ȡָ����RowList�б�
	 * @param startRowIndex �������ݼ���allRowList�����ʼ�±�
	 * @param length �������ݼ��ĳ���
	 * @param title ��������
	 * @param value ��������ֵ
	 * @return ����ָ��RowList�б�
	 * @throws ExcelIndexOutOfBoundsException ��ʼ�±겻��С�����������ֵ��length��ȡ�ĳ��Ȳ���С��0
	 * @throws ExcelIllegalArgumentException ָ�����ⲻ����
	 * @throws ExcelNoTitleException δ���ñ���
	 */
	/*public ArrayList<Row> getRowListByArg(int startRowIndex,int length,String title,String value) throws
				ExcelIndexOutOfBoundsException,ExcelIllegalArgumentException,ExcelNoTitleException{
		int count = rowList.size();
		ArrayList<Row> returnList = new ArrayList<>();
		
		int endIndex = isIndexOutOfBounds(count, startRowIndex, length);
		ArrayList<Row> rows = new ArrayList<>();
		for (int i = startRowIndex; i <= endIndex; i++) {
			rows.add(rowList.get(i));
		}
		returnList = getRowListByArg(rows, title, value);
		return returnList;
	}*/
	
	
	
	/**
	 * ��ȡָ��Row������CellList
	 * @param row ָ��Row
	 * @return ����cellList
	 * @throws ExcelNullParameterException ����Ϊnull
	 */
	/*public ArrayList<Cell> getOneRowAllCells(Row row) throws ExcelNullParameterException{
		if(row == null){
			throw new ExcelNullParameterException();
		}
		int count = row.getLastCellNum();
		if(count == -1){
			return new ArrayList<Cell>();
		}
		try {
			return getOneRowCellList(row, 0, count);
		} catch (ExcelIndexOutOfBoundsException e) {
			return new ArrayList<Cell>();
		}
	}*/
	
	/**
	 * ��ȡָ��Row��ָ����CellList
	 * @param row ָ��Row
	 * @param startColumnIndex ��ʼ�����±�ֵ
	 * @param length ���賤��ֵ
	 * @return ����cellList�����׳�����
	 * @throws ExcelIndexOutOfBoundsException ��ʼ�±겻��С�����������ֵ��length��ȡ�ĳ��Ȳ���С��0
	 * @throws ExcelNullParameterException ����rowΪnull
	 */
	public ArrayList<Cell> getOneRowCellList(Row row,int startColumnIndex,int length) throws 
					ExcelIndexOutOfBoundsException,ExcelNullParameterException{
		if(row == null){
			throw new ExcelNullParameterException();
		}
		ArrayList<Cell> cellList = new ArrayList<>();
		int count = row.getLastCellNum();
		if(count == -1){
			return cellList;
		}
		
		int endIndex = isIndexOutOfBounds(count, startColumnIndex, length);
		Cell c = null;
		for (int i = startColumnIndex; i <= endIndex; i++) {
			c = row.getCell(i);
			cellList.add(c);
		}
		return cellList;
		
	}
	
	/**
	 * ��ȡһ�е�����Cells
	 * @param columnIndex ָ�������±�ֵ
	 * @return ����cellList�����׳�����
	 * @throws ExcelIndexOutOfBoundsException �±�ֵԽ��
	 */
	/*public ArrayList<Cell> getOneColumnAllCells(int columnIndex) throws ExcelIndexOutOfBoundsException{
		int length = rowList.size();
		return getOneColumnCellList(columnIndex, 0, length);
	}*/
	
	/**
	 * ��ȡһ����ָ����Cells
	 * @param cloumnIndex ָ�������±�ֵ
	 * @param startRowIndex ��ʼ�����±�ֵ
	 * @param length ���賤��
	 * @return ����cellList�б�����׳�����
	 * @throws ExcelIndexOutOfBoundsException ��ʼ�±겻��С�����������ֵ��length��ȡ�ĳ��Ȳ���С��0
	 */
	/*public ArrayList<Cell> getOneColumnCellList(int cloumnIndex,int startRowIndex,int length) 
									throws ExcelIndexOutOfBoundsException{
		ArrayList<Cell> cellList = new ArrayList<>();
		int count = rowList.size();
		
		int endIndex = isIndexOutOfBounds(count, startRowIndex, length);
		Cell cell = null;
		Row row = null;
		for (int i = startRowIndex; i <= endIndex; i++) {
			row = rowList.get(i);
			if (row == null) {
				cellList.add(null);
			} else {
				cell = row.getCell(cloumnIndex);
				cellList.add(cell);
			}
		}
		return cellList;
	}*/
	
	/**
	 * �趨sheet�ı����б�
	 * @param rowIndex sheet�����±�
	 * @param startColumnIndex ��ʼ���±�ֵ
	 * @param length ���賤��
	 * @throws ExcelIndexOutOfBoundsException ��ʼ�±겻��С�����������ֵ��length��ȡ�ĳ��Ȳ���С��0
	 * @throws ExcelIllegalArgumentException ָ���±���Ϊnull,��û������
	 */
	public void setTitle(int rowIndex,int startColumnIndex,int length) throws ExcelIndexOutOfBoundsException,
													ExcelIllegalArgumentException{
		Row row = sheet.getRow(rowIndex);
		if(row == null)
			throw new ExcelIllegalArgumentException();
		int cellCount = row.getLastCellNum();
		if(cellCount == -1){
			throw new ExcelIllegalArgumentException();
		}
		titles.clear();
		
		ArrayList<Cell> cells = null;
		try {
			cells = getOneRowCellList(row, startColumnIndex, length);
		} catch (ExcelNullParameterException e) {
		}
		Cell cell = null;
		for (int i = 0; i < cells.size(); i++) {
			cell = cells.get(i);
			if (cell == null) {
				titles.put(startColumnIndex + i, "");
			} else {
				titles.put(startColumnIndex + i, getCellValue(cell));
			}
		}
	}
	
	/**
	 * ��ȡ�����б�
	 * @return ���ر���
	 */
	public HashMap<Integer, String> getTitle(){
		return titles;
	}
	
	/**
	 * ͨ������ֵ��ȡ�������±�
	 * @param titleName ָ������ֵ
	 * @return �������±�,�������򷵻�-1
	 * @throws ExcelNoTitleException δ���ñ���
	 */
	public int getTitleColIndexByValue(String titleName) throws ExcelNoTitleException{
		if(titles.isEmpty()){
			throw new ExcelNoTitleException();
		}
		for(Map.Entry<Integer, String> entry : titles.entrySet()){
			if(entry.getValue().equals(titleName)) 
				return entry.getKey();
		}
		return -1;
	}
	
	/**
	 * ��ȡָ��Cell������ֵ
	 * @param cell ָ��Cell
	 * @return ������������String���ͷ���;
	 * <br>�������͸�ʽ(yyyy-MM-dd,yyyy-MM,MM-dd),��Ԫ��ʽΪDATE_NUMʱֻ����(yyyy-MM-dd);
	 * <br>��ʽ���͸�ʽ���ع�ʽ���㷽��������ֵ;
	 * <br>cellΪnull,��û��ֵ,��ȡֵ�����򷵻�""��
	 */
	public static String getCellValue(Cell cell){
		if(cell == null){
			return "";
		}
		String string = "";
		try {
			TypeEnum cellType = CellTypeUtil.getCellType(cell);
			
			switch (cellType) {
			case STRING:
				string =  cell.getStringCellValue().trim();
				break;
			case NUMERIC:
				double dd = cell.getNumericCellValue();
				int in = (int) dd;
				double last = dd - in;
				if (last == 0) // doubleΪ����
					string = String.valueOf(in);
				else
					string = String.valueOf(dd);
				break;
			case DATE_NUM:
				double d = cell.getNumericCellValue();
				string = CellTypeUtil.getFormatDate(d);
				break;
			case DATE_STR:
				string =  cell.getStringCellValue().trim();
				string = CellTypeUtil.getFormatDate(string);
				break;
			case ERROR:
				string = String.valueOf(cell.getErrorCellValue());
				break;
			case FORMULA:
				string = cell.getCellFormula();
				break;
			case BOOLEAN:
				string = String.valueOf(cell.getBooleanCellValue());
				break;
			case BLANK:
				string = "";
				break;
			default:
				string = "";
			}
			//ȥ����βȫ�ǿհ׷�
			while (string.startsWith("��")) {
				string = string.substring(1, string.length()).trim();
			}
			while (string.endsWith("��")) {
				string = string.substring(0, string.length() - 1).trim();
			}
			/*if (string != null) {//ȥ�����С��س����Ʊ��  
		        Pattern p = Pattern.compile("\\s*|\t|\r|\n");  
		        Matcher m = p.matcher(string);  
		        string = m.replaceAll("");  
		    }  */
			return string.trim();
		} catch (Exception e) {//δ֪����
			return "";
		}
	}
	
	/**
	 * ��ȡָ��Cell������ֵ
	 * @param cell ָ��Cell
	 * @param evaluator ���ڶ�ȡ��ʽֵ��FormulaEvaluatorʵ��
	 * @return ������������String���ͷ���;
	 * <br>�������͸�ʽ(yyyy-MM-dd,yyyy-MM,MM-dd),��Ԫ��ʽΪDATE_NUMʱֻ����(yyyy-MM-dd);
	 * <br>��ʽ���͸�ʽ���ع�ʽ����ֵ(��evaluatorΪnull�򷵻ع�ʽ);
	 * <br>cellΪnull,��û��ֵ,��ȡֵ�����򷵻�""��
	 */
	public static String getCellValue(Cell cell,FormulaEvaluator evaluator){
		if(cell == null){
			return "";
		}
		String string = "";
		try {
			TypeEnum cellType = CellTypeUtil.getCellType(cell);
			
			switch (cellType) {
			case FORMULA:
				if(evaluator == null){
					string = cell.getCellFormula();
				}else{
					string = getFormulaValue(evaluator.evaluate(cell));
				}
				break;
			default:
				string = getCellValue(cell);
			}
			//ȥ����βȫ�ǿհ׷�
			while (string.startsWith("��")) {
				string = string.substring(1, string.length()).trim();
			}
			while (string.endsWith("��")) {
				string = string.substring(0, string.length() - 1).trim();
			}
			
			return string.trim();
		} catch (Exception e) {//δ֪����
			return "";
		}
	}
	
	
	/**
	 * ��ȡ�ϲ���Ԫ���ֵ 
	 * @param cell ָ����Ԫ��
	 * @return ��Ԫ���ֵ
	 * @rhrows ExcelNullParameterException ����cellΪnull
	 * @throws ExcelIllegalArgumentException cell���Ǻϲ���Ԫ��
	 */
	public static String getCellValueOfMergedRegion(Sheet sheet,Cell cell) throws ExcelNullParameterException,ExcelIllegalArgumentException{  
		if(cell == null){
			throw new ExcelNullParameterException();
		}
		
		int result = isCellInMergedRegion(sheet,cell);
		if (result == 2) { // ��Ԫ���ںϲ������ڲ�
			int sheetMergeCount = sheet.getNumMergedRegions();
			int rowIndex = cell.getRowIndex();
			int columnIndex = cell.getColumnIndex();
			for (int i = 0; i < sheetMergeCount; i++) {
				CellRangeAddress ca = sheet.getMergedRegion(i);
				int firstColumn = ca.getFirstColumn();
				int lastColumn = ca.getLastColumn();
				int firstRow = ca.getFirstRow();
				int lastRow = ca.getLastRow();

				if (rowIndex >= firstRow && rowIndex <= lastRow) {
					if (columnIndex >= firstColumn && columnIndex <= lastColumn) { // ȷ������λ��
						Row fRow = sheet.getRow(firstRow);
						Cell fCell = fRow.getCell(firstColumn);
						return getCellValue(fCell); // ���ظ�����ĵ�һ����Ԫ���ֵ
					}
				}
			}
			return "";
		} else if (result == 1) { // ��Ԫ��Ϊ�ϲ�����ĵ�һ����Ԫ
			return getCellValue(cell); // ֱ�ӷ��ظõ�Ԫ���ֵ
		} else {
			throw new ExcelIllegalArgumentException();
		}
		
	}    
	
	/**
	 * ǿ�ƽ�Numeric���͵�ֵת��Ϊ����
	 * @param numeric numeric��ʽ�ĵ�Ԫ����ֵ
	 * @return yyyy-MM-dd��ʽ������
	 * @throws NullPointerException ����ǿת����
	 */
	public static String changeNumericToDate(double numeric) throws NullPointerException{
		String string = "";
		Date date = DateUtil.getJavaDate(numeric);
		if(date == null){
			date = HSSFDateUtil.getJavaDate(numeric);
		}
		string = new SimpleDateFormat("yyyy-MM-dd").format(date);
		return string;
	}
	
	/**  
	* �ж�sheetҳ���Ƿ��кϲ���Ԫ��   
	* @param sheet   
	* @return  �кϲ���Ԫ�񷵻�true���򷵻�false
	*/  
	public static boolean hasMerged(Sheet sheet) {  
	     return sheet.getNumMergedRegions() > 0 ? true : false;  
	} 
	
	/**
	 * �ж�ָ���ĵ�Ԫ���Ƿ��Ǻϲ���Ԫ��  
	 * @param cell ָ����Ԫ��
	 * @return 1(��Ԫ���Ǻϲ�����ĵ�һ����Ԫ)��2(�ϲ������ڲ��ĵ�Ԫ)��-1(���Ǻϲ������ڵĵ�Ԫ)
	 * @throws ExcelNullParameterException ����Ϊnull
	 */
	public static int isCellInMergedRegion(Sheet sheet,Cell cell) throws ExcelNullParameterException{
		if (hasMerged(sheet)) {
			if(cell == null){
				throw new ExcelNullParameterException();
			}
			int sheetMergeCount = sheet.getNumMergedRegions();
			int rowIndex = cell.getRowIndex();
			int columnIndex = cell.getColumnIndex();
			for (int i = 0; i < sheetMergeCount; i++) {
				CellRangeAddress range = sheet.getMergedRegion(i);
				int firstColumn = range.getFirstColumn();
				int lastColumn = range.getLastColumn();
				int firstRow = range.getFirstRow();
				int lastRow = range.getLastRow();
				if(rowIndex == firstRow && columnIndex == firstColumn){//��Ԫ��Ϊ�Ǻϲ�����ĵ�һ����Ԫ
					return 1;
				}else if (rowIndex >= firstRow && rowIndex <= lastRow) {//��Ԫ���ںϲ������ڲ�
					if (columnIndex >= firstColumn && columnIndex <= lastColumn) {
						return 2;
					}
				}
			}
			return -1;
		}else{
			return -1;
		}
	}
	
	
	/**
	 * �ж�ָ���ĵ�Ԫ���Ƿ��Ǻϲ���Ԫ��  
	 * @param rowIndex ��Ԫ�����±�
	 * @param columnIndex ��Ԫ�����±�
	 * @return 1(��Ԫ���Ǻϲ�����ĵ�һ����Ԫ)��2(�ϲ������ڲ��ĵ�Ԫ)��-1(���Ǻϲ������ڵĵ�Ԫ)
	 * @throws ExcelIndexOutOfBoundsException ����С����
	 */
	public static int isCellInMergedRegion(Sheet sheet,int rowIndex,int columnIndex) throws ExcelIndexOutOfBoundsException{
		if (hasMerged(sheet)) {
			if(rowIndex <0 || columnIndex <0){
				throw new ExcelIndexOutOfBoundsException();
			}
			int sheetMergeCount = sheet.getNumMergedRegions();
			for (int i = 0; i < sheetMergeCount; i++) {
				CellRangeAddress range = sheet.getMergedRegion(i);
				int firstColumn = range.getFirstColumn();
				int lastColumn = range.getLastColumn();
				int firstRow = range.getFirstRow();
				int lastRow = range.getLastRow();
				if(rowIndex == firstRow && columnIndex == firstColumn){//��Ԫ��Ϊ�Ǻϲ�����ĵ�һ����Ԫ
					return 1;
				}else if (rowIndex >= firstRow && rowIndex <= lastRow) {//��Ԫ���ںϲ������ڲ�
					if (columnIndex >= firstColumn && columnIndex <= lastColumn) {
						return 2;
					}
				}
			}
			return -1;//��Ԫ���ںϲ�������
		}else{
			return -1;
		}
	}	
	
	/**
	 * �����ͷ��ع�ʽ����ֵ
	 * @param cellValue 
	 * @return ��ֵ���ı����ͷ��ؾ���ֵ����������""��
	 */
	private static String getFormulaValue(CellValue cellValue) {
        String value = "";
        switch (cellValue.getCellTypeEnum()) {
        case STRING:
            value = cellValue.getStringValue();
            break;
        case NUMERIC:
            value = String.valueOf(cellValue.getNumberValue());
            break;
        default:
            break;
        }
        return value;
    }
	
	/**
	 * �ж����ݳ��ȡ���ʼ�±�Ͷ�ȡ���Ȳ����Ƿ�Խ��
	 * @param count �����ܳ���,����С��1
	 * @param startIndex ��ʼ�±겻��С�����������ֵ
	 * @param length ��ȡ�ĳ���,����С��0
	 * @return �������Խ���׳��쳣,���򷵻�Ҫ��ȡ�����һ���±�ֵ(�����ȡ���ȴ����ܳ���,�򷵻�����±�ֵ)
	 * @throws ExcelIndexOutOfBoundsException ����Խ�����
	 */
	protected int isIndexOutOfBounds(int count,int startIndex,int length) throws ExcelIndexOutOfBoundsException{
		if(count<1){ //���ݳ���С��1
			throw new ExcelIndexOutOfBoundsException();
		}
		if(length<0){//��ȡ����С����
			throw new ExcelIndexOutOfBoundsException();
        }
        if(startIndex > count -1 || startIndex < 0){//��ʼ�±���������±�ֵ��С����
        	throw new ExcelIndexOutOfBoundsException();
        }
        //Ҫ��ȡ�����һ���±�,����±�Խ�磬���ȡ�����һ��ֵ
        int endIndex = startIndex + length - 1;
		if (endIndex >= count)
			endIndex = count - 1;
        return endIndex;
	}
	
}
