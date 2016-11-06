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

	//%%%%%%%%-------�ֶβ��� ��ʼ----------%%%%%%%%%
	/**
	 * Excel�ļ�·��
	 */
	private String excelPath = "";
	
	/**
	 * Excel�ļ�
	 */
	private File excelFile;
	
	/**
	 *�ļ���׺��(xls,xlsx) 
	 */
	private String extension = "";
	
	/**
	 * ����Excel��Workbook����
	 */
	private Workbook workbook = null;
	
	/**
	 * ��־�������
	 */
	private static Logger logger = Logger.getLogger("excelLog");
	
	/**
	 * ���һ�ζ�ȡ��sheet�б�
	 */
	public ArrayList<SheetReadBean> sheetList = null; 

	
	/**
	 * Ĭ��Excel���ݵĿ�ʼ�Ƚ���λ��Ϊ��һ�У�����ֵΪ0��
	 */
	private int compareColumnPos = 0;

	/**
	 * ���ļ��ϲ�ʱ���������ظ�ʱ�Ƿ���и��ǣ�
	 * Ĭ��Ϊtrue
	 */
	private boolean isOverWrite = true;
	
	/**
	 * ���ļ��ϲ�ʱ�Ƿ���Ҫ�����ݱȽ�(����ͬ�����ݲ��ظ�����),Ĭ��ֵΪtrue
	 * (��������дĿ����������Ч����isOverWrite=falseʱ��Ч)
	 */
	private boolean isNeedCompare = true;
	

	
	
	//%%%%%%%%-------�ֶβ��� ����----------%%%%%%%%%
	

	/**
	 * @param excelPath  �ļ�·��
	 * @throws IllegalArgumentException �ļ������ڻ��ʽ����
	 * @throws NullPointerException �ļ�·��Ϊnull
	 * @throws SecurityException �ļ��ܾ�����
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
				//���ַ�ʽ Excel 2003/2007/2010 ���ǿ��Դ����  
		        this.workbook = WorkbookFactory.create(is) ;
			}else {
				throw new IllegalArgumentException("�ļ�����Excel�ļ�");
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
	 * ����newһ���µĶ��󲢷���
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
	 * ����ļ��Ƿ�ΪExcel�ļ�
	 * @param filePath �ļ�·��
	 * @return ����ļ�ΪExcel��ʽ�򷵻�true,����false
	 * @throws IllegalArgumentException �ļ�������
	 * @throws NullPointerException �ļ�·��Ϊnull
	 * @throws SecurityException �ļ��ܾ�����
	 */
	public static boolean isExcelFile(String filePath) throws IllegalArgumentException,
															NullPointerException,SecurityException{
		try {
			File file = new File(filePath);
			if(!file.exists()){
				throw new IllegalArgumentException("·������,�ļ�������");
			}else{
				String name = file.getName();
				String ext = name.substring(name.lastIndexOf("."));
				if(ext.equals(".xls") || ext.equals(".xlsx")) return true;
				else return false;
			}
		} catch (NullPointerException e) {
			throw new NullPointerException("�ļ�·��Ϊ��");
		} catch (SecurityException  e) {	
			throw new SecurityException("�ļ��ܾ�����");
		}
	}
	
	/**
	 * ��ȡ��һ��sheet
	 */
	public void readFirstSheet(){
		this.sheetList = changeSL2SBL(getSheetList(0, 1));
		readAllRows(sheetList.get(0));
	}
	
	/**
	 * ��ȡ�±�ΪIndex��Sheet
	 * @param index sheet���±�ֵ
	 * @throws IndexOutOfBoundsException ����Խ�����
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
	 * ͨ�����ƶ�ȡsheet
	 * @param name sheet������
	 * @throws IllegalArgumentException ���ƴ����޷���ȡָ��sheet
	 */
	public void readSheetByName(String name) throws IllegalArgumentException{
		Sheet sheet = workbook.getSheet(name);
		if(sheet != null){
			sheetList.add(new SheetReadBean(sheet));
			readAllRows(sheetList.get(0));
		}
		else
			throw new IllegalArgumentException("�޷���ȡָ�����Ƶ�sheet");
	}
	
	/**
	 * ��ȡָ����Χ��sheet�б�
	 * @param startIndex sheet��ʼ���±�ֵ
	 * @param length Ҫ��ȡsheets�ĳ���
	 * @throws IndexOutOfBoundsException ����Խ�����
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
	 * ��ȡ���е�sheet
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
	 * ��ȡ��Ҫ��sheet�б�
	 * @param startIndex sheet��ʼ���±�ֵ
	 * @param length Ҫ��ȡsheets�ĳ���,������ȹ������ȡ�����ݽ�β
	 * @return Sheet �б�
	 * @throws IndexOutOfBoundsException ��������
	 */
	public ArrayList<Sheet> getSheetList(int startIndex,int length) throws IndexOutOfBoundsException{
		ArrayList<Sheet> sheetList = new ArrayList<>();
        int sheetCount = workbook.getNumberOfSheets();  //Sheet������  
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
	 * ��SheetListת��ΪSheetBeanList
	 * @param sheetList
	 * @return sheetBean�б�
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
	 * ��ȡSheetBean�����е���
	 * @param sheetBean Ҫ��ȡ��sheetBean
	 */
	private void readAllRows(SheetReadBean sheetBean){
		Sheet sheet = sheetBean.getSheet();
		int rowsCount = sheet.getLastRowNum() + 1;
		readRows(sheetBean, 0, rowsCount);
	}

	/**
	 * ��ȡsheetBean��ָ������
	 * @param sheetBean Ҫ��ȡ��sheetBean
	 * @param startIndex ��ʼ�����±�
	 * @param length ��ȡ����
	 * @throws IndexOutOfBoundsException ����Խ�����
	 */
	private void readRows(SheetReadBean sheetBean,int startIndex,int length) throws IndexOutOfBoundsException{
		Sheet sheet = sheetBean.getSheet();
		int rowsCount = sheet.getLastRowNum() + 1;
		try {
			int endIndex = isIndexOutOfBounds(rowsCount, startIndex, length);
			for(int i = startIndex;i<=endIndex;i++){
				if(endIndex == 0){//ֻ��sheet�ĵ�һ��
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
	 * �ж����ݳ��ȡ���ʼ�±�Ͷ�ȡ���Ȳ����Ƿ�Խ��
	 * @param count �����ܳ���,����С��1
	 * @param startIndex ��ʼ�±겻��С�����������ֵ
	 * @param length ��ȡ�ĳ���,����С��0
	 * @return �������Խ���׳��쳣,���򷵻�Ҫ��ȡ�����һ���±�ֵ(����±�Խ��,�򷵻�����±�ֵ)
	 * @throws IndexOutOfBoundsException 
	 */
	public static int isIndexOutOfBounds(int count,int startIndex,int length) throws IndexOutOfBoundsException{
		if(count<1){
			throw new IndexOutOfBoundsException("���ݳ���С��1");
		}
		if(length<0){
        	throw new IndexOutOfBoundsException("��ȡ����С����");
        }
        if(startIndex > count -1 || startIndex < 0){
        	throw new IndexOutOfBoundsException("��ʼ�±���������±�ֵ��С����");
        }
        //Ҫ��ȡ�����һ���±�,����±�Խ�磬���ȡ�����һ��ֵ
        int endIndex = startIndex + length - 1;
		if (endIndex >= count)
			endIndex = count - 1;
        return endIndex;
	}
	
	



	
	/**
	 * �Զ������ļ���չ�������ö�Ӧ��д�뷽��
	 * 
	 * @Title: writeExcel
	 * @Date : 2014-9-11 ����01:50:38
	 * @param rowList
	 * @param xlsPath
	 * @throws IOException
	 */
	public void writeExcel(List<Row> rowList, String xlsPath) throws IOException {

		//��չ��Ϊ��ʱ��
		if (xlsPath.equals("")){
			throw new IOException("�ļ�·������Ϊ�գ�");
		}
		
		//��ȡ��չ��
		//String ext = xlsPath.substring(xlsPath.lastIndexOf(".")+1);
		
		/*try {
			
			if("xls".equals(ext)){				//ʹ��xls��ʽд��
				//writeExcel_xls(rowList,xlsPath);
			}else if("xls".equals(ext)){		//ʹ��xlsx��ʽд��
				//writeExcel_xlsx(rowList,xlsPath);
			}else{									//���γ���xls��xlsx��ʽд��
				//out("��Ҫ�������ļ�û����չ�������ڳ�����xls��ʽд��...");
				try{
					//writeExcel_xls(rowList,xlsPath);
				} catch (IOException e1) {
					//out("������xls��ʽд�룬���ʧ�ܣ������ڳ�����xlsx��ʽ��ȡ...");
					try{
						//writeExcel_xlsx(rowList,xlsPath);
					} catch (IOException e2) {
						//out("������xls��ʽд�룬���ʧ�ܣ�\n����ȷ�������ļ���Excel�ļ�����������Ȼ�����ԡ�");
						throw e2;
					}
				}
			}
		} catch (IOException e) {
			throw e;
		}*/
	}


	/**
	 * �޸�Excel��97-03�棬xls��ʽ��
	 * 
	 * @Title: writeExcel_xls
	 * @Date : 2014-9-11 ����01:50:38
	 * @param rowList
	 * @param src_xlsPath
	 * @param dist_xlsPath
	 * @throws IOException
	 */
	/*public void writeExcel_xls(List<Row> rowList, String src_xlsPath, String dist_xlsPath) throws IOException {

		// �ж��ļ�·���Ƿ�Ϊ��
		if (dist_xlsPath == null || dist_xlsPath.equals("")) {
			//out("�ļ�·������Ϊ��");
			throw new IOException("�ļ�·������Ϊ��");
		}
		// �ж��ļ�·���Ƿ�Ϊ��
		if (src_xlsPath == null || src_xlsPath.equals("")) {
			//out("�ļ�·������Ϊ��");
			throw new IOException("�ļ�·������Ϊ��");
		}

		// �ж��б��Ƿ������ݣ����û�����ݣ��򷵻�
		if (rowList == null || rowList.size() == 0) {
			//out("�ĵ�Ϊ��");
			return;
		}

		try {
			HSSFWorkbook wb = null;

			// �ж��ļ��Ƿ����
			File file = new File(dist_xlsPath);
			if (file.exists()) {
				// �����д����ɾ����
				if (isOverWrite) {
					file.delete();
					// ����ļ������ڣ��򴴽�һ���µ�Excel
					// wb = new HSSFWorkbook();
					// wb.createSheet("Sheet1");
					wb = new HSSFWorkbook(new FileInputStream(src_xlsPath));
				} else {
					// ����ļ����ڣ����ȡExcel
					wb = new HSSFWorkbook(new FileInputStream(file));
				}
			} else {
				// ����ļ������ڣ��򴴽�һ���µ�Excel
				// wb = new HSSFWorkbook();
				// wb.createSheet("Sheet1");
				wb = new HSSFWorkbook(new FileInputStream(src_xlsPath));
			}

			// ��rowlist������д��Excel��
			writeExcel(wb, rowList, dist_xlsPath);

		} catch (IOException e) {
			e.printStackTrace();
		}
	}
*/


	/**
	 * �޸�Excel��2007�棬xlsx��ʽ��
	 * 
	 * @Title: writeExcel_xlsx
	 * @Date : 2014-9-11 ����01:50:38
	 * @param rowList
	 * @param xlsPath
	 * @throws IOException
	 */
	/*public void writeExcel_xlsx(List<Row> rowList, String src_xlsPath, String dist_xlsPath) throws IOException {

		// �ж��ļ�·���Ƿ�Ϊ��
		if (dist_xlsPath == null || dist_xlsPath.equals("")) {
			//out("�ļ�·������Ϊ��");
			throw new IOException("�ļ�·������Ϊ��");
		}
		// �ж��ļ�·���Ƿ�Ϊ��
		if (src_xlsPath == null || src_xlsPath.equals("")) {
			//out("�ļ�·������Ϊ��");
			throw new IOException("�ļ�·������Ϊ��");
		}

		// �ж��б��Ƿ������ݣ����û�����ݣ��򷵻�
		if (rowList == null || rowList.size() == 0) {
			//out("�ĵ�Ϊ��");
			return;
		}

		try {
			// ��ȡ�ĵ�
			XSSFWorkbook wb = null;

			// �ж��ļ��Ƿ����
			File file = new File(dist_xlsPath);
			if (file.exists()) {
				// �����д����ɾ����
				if (isOverWrite) {
					file.delete();
					// ����ļ������ڣ��򴴽�һ���µ�Excel
					// wb = new XSSFWorkbook();
					// wb.createSheet("Sheet1");
					wb = new XSSFWorkbook(new FileInputStream(src_xlsPath));
				} else {
					// ����ļ����ڣ����ȡExcel
					wb = new XSSFWorkbook(new FileInputStream(file));
				}
			} else {
				// ����ļ������ڣ��򴴽�һ���µ�Excel
				// wb = new XSSFWorkbook();
				// wb.createSheet("Sheet1");
				wb = new XSSFWorkbook(new FileInputStream(src_xlsPath));
			}
			// ��rowlist��������ӵ�Excel��
			writeExcel(wb, rowList, dist_xlsPath);

		} catch (IOException e) {
			e.printStackTrace();
		}
	}
*/


	

	
	/**
	 * �޸�Excel�������Ϊ
	 * 
	 * @Title: WriteExcel
	 * @Date : 2014-9-11 ����01:33:59
	 * @param wb
	 * @param rowList
	 * @param xlsPath
	 */
	/*private void writeExcel(Workbook wb, List<Row> rowList, String xlsPath) {

		if (wb == null) {
			//out("�����ĵ�����Ϊ�գ�");
			return;
		}

		Sheet sheet = wb.getSheetAt(0);// �޸ĵ�һ��sheet�е�ֵ

		// ���ÿ����д����ô��ӿ�ʼ��ȡ��λ��д���������ȡԴ�ļ����µ��С�
		int lastRowNum = isOverWrite ? startReadRowPos : sheet.getLastRowNum() + 1;
		int t = 0;//��¼������ӵ�����
		//out("Ҫ��ӵ�����������Ϊ��"+rowList.size());
		for (Row row : rowList) {
			if (row == null) continue;
			// �ж��Ƿ��Ѿ����ڸ�����
			int pos = findInExcel(sheet, row);

			Row r = null;// ����������Ѿ����ڣ����ȡ����д�������Զ��������С�
			if (pos >= 0) {
				sheet.removeRow(sheet.getRow(pos));
				r = sheet.createRow(pos);
			} else {
				r = sheet.createRow(lastRowNum + t++);
			}
			
			//�����趨��Ԫ����ʽ
			CellStyle newstyle = wb.createCellStyle();
			
			//ѭ��Ϊ���д�����Ԫ��
			for (int i = row.getFirstCellNum(); i < row.getLastCellNum(); i++) {
				Cell cell = r.createCell(i);// ��ȡ��������
				cell.setCellValue(getCellValue(row.getCell(i)));// ���Ƶ�Ԫ���ֵ���µĵ�Ԫ��
				// cell.setCellStyle(row.getCell(i).getCellStyle());//����
				if (row.getCell(i) == null) continue;
				copyCellStyle(row.getCell(i).getCellStyle(), newstyle); // ��ȡԭ���ĵ�Ԫ����ʽ
				cell.setCellStyle(newstyle);// ������ʽ
				// sheet.autoSizeColumn(i);//�Զ���ת�п��
			}
		}
		//out("���м�⵽�ظ�����Ϊ:" + (rowList.size() - t) + " ��׷������Ϊ��"+t);
		
		// ͳһ�趨�ϲ���Ԫ��
		setMergedRegion(sheet);
		
		try {
			// ���½�����д��Excel��
			FileOutputStream outputStream = new FileOutputStream(xlsPath);
			wb.write(outputStream);
			outputStream.flush();
			outputStream.close();
		} catch (Exception e) {
			//out("д��Excelʱ�������� ");
			e.printStackTrace();
		}
	}*/

	/**
	 * ����ĳ�������Ƿ���Excel���д��ڣ�����������
	 * 
	 * @Title: findInExcel
	 * @Date : 2014-9-11 ����02:23:12
	 * @param sheet
	 * @param row
	 * @return
	 */
	/*private int findInExcel(Sheet sheet, Row row) {
		int pos = -1;

		try {
			// �����дĿ���ļ������߲���Ҫ�Ƚϣ���ֱ�ӷ���
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
	 * ����һ����Ԫ����ʽ��Ŀ�ĵ�Ԫ����ʽ
	 * 
	 * @param fromStyle
	 * @param toStyle
	 */
	/*public static void copyCellStyle(CellStyle fromStyle, CellStyle toStyle) {
		toStyle.setAlignment(fromStyle.getAlignment());
		// �߿�ͱ߿���ɫ
		toStyle.setBorderBottom(fromStyle.getBorderBottom());
		toStyle.setBorderLeft(fromStyle.getBorderLeft());
		toStyle.setBorderRight(fromStyle.getBorderRight());
		toStyle.setBorderTop(fromStyle.getBorderTop());
		toStyle.setTopBorderColor(fromStyle.getTopBorderColor());
		toStyle.setBottomBorderColor(fromStyle.getBottomBorderColor());
		toStyle.setRightBorderColor(fromStyle.getRightBorderColor());
		toStyle.setLeftBorderColor(fromStyle.getLeftBorderColor());

		// ������ǰ��
		toStyle.setFillBackgroundColor(fromStyle.getFillBackgroundColor());
		toStyle.setFillForegroundColor(fromStyle.getFillForegroundColor());

		// ���ݸ�ʽ
		toStyle.setDataFormat(fromStyle.getDataFormat());
		toStyle.setFillPattern(fromStyle.getFillPattern());
		// toStyle.setFont(fromStyle.getFont(null));
		toStyle.setHidden(fromStyle.getHidden());
		toStyle.setIndention(fromStyle.getIndention());// ��������
		toStyle.setLocked(fromStyle.getLocked());
		toStyle.setRotation(fromStyle.getRotation());// ��ת
		toStyle.setVerticalAlignment(fromStyle.getVerticalAlignment());
		toStyle.setWrapText(fromStyle.getWrapText());

	}
*/
	/**
	 * ��ȡ�ϲ���Ԫ���ֵ
	 * 
	 * @param sheet
	 * @param row
	 * @param column
	 * @return
	 */
/*	public void setMergedRegion(Sheet sheet) {
		int sheetMergeCount = sheet.getNumMergedRegions();

		for (int i = 0; i < sheetMergeCount; i++) {
			// ��ȡ�ϲ���Ԫ��λ��
			CellRangeAddress ca = sheet.getMergedRegion(i);
			int firstRow = ca.getFirstRow();
			if (startReadRowPos - 1 > firstRow) {// �����һ���ϲ���Ԫ���ʽ����ʽ���ݵ����棬��������
				continue;
			}
			int lastRow = ca.getLastRow();
			int mergeRows = lastRow - firstRow;// �ϲ�������
			int firstColumn = ca.getFirstColumn();
			int lastColumn = ca.getLastColumn();
			// ���ݺϲ��ĵ�Ԫ��λ�úʹ�С���������е������и�ʽ��
			for (int j = lastRow + 1; j <= sheet.getLastRowNum(); j++) {
				// �趨�ϲ���Ԫ��
				sheet.addMergedRegion(new CellRangeAddress(j, j + mergeRows, firstColumn, lastColumn));
				j = j + mergeRows;// �����Ѻϲ�����
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