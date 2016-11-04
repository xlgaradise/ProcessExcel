package testExcel;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.List;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;


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
	 * ��ʼ��ȡ����λ��Ϊ��һ�У�����ֵΪ0��
	 */
	private int startReadRowPos = 0;

	/**
	 * ������ȡ����λ��Ϊ���һ��,Ĭ��ֵΪ-1���ø�������ʾ������n�У�
	 */
	private int endReadRowPos = -1;
	
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
	
	/**
	 * �趨�Ƿ�ֻ������һ��sheet
	 */
	private boolean onlyReadOneSheet = true;
	
	/**
	 * �趨������sheet������ֵ,Ĭ��ֵΪ0
	 */
	private int selectedSheetIndex = 0;
	
	/**
	 * �趨������sheet������
	 */
	private String selectedSheetName = "";
	
	/**
	 * �趨��ʼ��ȡ��sheet��Ĭ��Ϊ0
	 */
	private int startSheetIdx = 0;

	/**
	 * �趨������ȡ��sheet��Ĭ��Ϊ-1���ø�������ʾ������n��	
	 */
	private int endSheetIdx = -1;
	
	/**
	 * �趨�Ƿ��ӡ��Ϣ
	 */
	private boolean printMsg = true;
	
	
	//%%%%%%%%-------�ֶβ��� ����----------%%%%%%%%%
	

	/**
	 * @param excelPath  �ļ�·��
	 * @throws IllegalArgumentException �ļ������ڻ��߸�ʽ��ΪExcel�ļ�
	 * @throws NullPointerException �ļ�·��Ϊnull
	 * @throws SecurityException �ļ��ܾ�����
	 */
	public ExcelUtil(String excelPath) throws IllegalArgumentException,
											NullPointerException,SecurityException{
		try {
			if(isExcelFile(excelPath)){
				this.excelPath = excelPath;
				this.excelFile = new File(excelPath);
				String name = this.excelFile.getName();
				this.extension = name.substring(name.lastIndexOf("."));
			}
		} catch (IllegalArgumentException e) {
			throw e;
		}catch (NullPointerException e) {
			throw e;
		}catch (SecurityException e) {
			throw e;
		}
	}
	
	/**
	 * ����newһ���µĶ��󲢷���
	 * @return
	 */
	public ExcelUtil returnNewInstance(){
		ExcelUtil instance = new  ExcelUtil(this.excelPath);
		return instance;
	}
	
	/**
	 * ����ļ��Ƿ�ΪExcel�ļ�
	 * @param filePath �ļ�·��
	 * @return ����ļ�ΪExcel��ʽ�򷵻�true,����false
	 * @throws IllegalArgumentException �ļ������ڻ��߸�ʽ��ΪExcel�ļ�
	 * @throws NullPointerException �ļ�·��Ϊnull
	 * @throws SecurityException �ļ��ܾ�����
	 */
	public static boolean isExcelFile(String filePath) throws IllegalArgumentException,
															NullPointerException,SecurityException{
		try {
			File file = new File(filePath);
			if(!file.exists()){
				saveErrorMessage("isExcelFile", "["+filePath+"]file is not exist.");
				throw new IllegalArgumentException("·������,�ļ�������");
			}else{
				String name = file.getName();
				String ext = name.substring(name.lastIndexOf("."));
				if(ext.equals("xls") || ext.equals("xlsx")) return true;
				else return false;
			}
		} catch (NullPointerException e) {
			saveErrorMessage("isExcelFile", "["+filePath+"]filePath is null.");
			throw new NullPointerException("�ļ�·��Ϊ��");
		} catch (SecurityException  e) {	
			saveErrorMessage("isExcelFile", "["+filePath+"]file is denies read.");
			throw new SecurityException("�ļ��ܾ�����");
		}
	}
	
	/**
	 * �Զ������ļ���չ�������ö�Ӧ�Ķ�ȡ����
	 * @return 
	 */
	public List<Row> readExcel() throws IOException{
		try {
			if (extension.equals("xls")) { // ʹ��xls��ʽ��ȡ
				return readExcel_xls();
			}else if("xlsx".equals(extension)) { // ʹ��xlsx��ʽ��ȡ
				return readExcel_xlsx();
			}else{
				return null;
			}
		} catch (IOException e) {
			throw e;
			
		}
	}
	
	
	/***
	 * ��ȡExcel(97-03�棬xls��ʽ)
	 * 
	 * @throws Exception
	 * 
	 * @Title: readExcel
	 * @Date : 2014-9-11 ����09:53:21
	 */
	private List<Row> readExcel_xls() throws IOException {
		HSSFWorkbook wb = null;// ����Workbook���Ĳ�����������ɾ��Excel
		List<Row> rowList = new ArrayList<Row>();

		try {
			// ��ȡExcel
			wb = new HSSFWorkbook(new FileInputStream(excelFile));

			// ��ȡExcel 97-03�棬xls��ʽ
			rowList = readExcel(wb);

		} catch (IOException e) {
			e.printStackTrace();
		}
		return rowList;
	}

	
	/**
	 * //��ȡExcel 2007�棬xlsx��ʽ
	 * 
	 * @Title: readExcel_xlsx
	 * @Date : 2014-9-11 ����11:43:11
	 * @return
	 * @throws Exception
	 */
	private List<Row> readExcel_xlsx() throws IOException {
		

		XSSFWorkbook wb = null;
		List<Row> rowList = new ArrayList<Row>();
		try {
			FileInputStream fis = new FileInputStream(excelFile);
			// ȥ��Excel
			wb = new XSSFWorkbook(fis);

			// ��ȡExcel 2007�棬xlsx��ʽ
			rowList = readExcel(wb);

		} catch (IOException e) {
			e.printStackTrace();
		}
		return rowList;
	}
	
	
	/**
	 * ͨ�ö�ȡExcel
	 * @param wb
	 * @return
	 */
	private List<Row> readExcel(Workbook wb) {
		List<Row> rowList = new ArrayList<Row>();
		
		int sheetCount = 1;//��Ҫ������sheet���� 
		
		Sheet sheet = null;
		if(onlyReadOneSheet){	//ֻ����һ��sheet
			// ��ȡ�趨������sheet(����趨�����ƣ������Ʋ飬��������ֵ��)
			sheet =selectedSheetName.equals("")? wb.getSheetAt(selectedSheetIndex):wb.getSheet(selectedSheetName);
		}else{							//�������sheet
			sheetCount = wb.getNumberOfSheets();//��ȡ���Բ�����������
		}
		
		// ��ȡsheet��Ŀ
		for(int t=startSheetIdx; t<sheetCount+endSheetIdx;t++){
			// ��ȡ�趨������sheet
			if(!onlyReadOneSheet) {
				sheet =wb.getSheetAt(t);
			}
			
			//��ȡ����к�
			int lastRowNum = sheet.getLastRowNum();

			if(lastRowNum>0){	//���>0����ʾ������
				out("\n��ʼ��ȡ��Ϊ��"+sheet.getSheetName()+"�������ݣ�");
			}
			
			Row row = null;
			// ѭ����ȡ
			for (int i = startReadRowPos; i <= lastRowNum + endReadRowPos; i++) {
				row = sheet.getRow(i);
				if (row != null) {
					rowList.add(row);
					out("��"+(i+1)+"�У�",false);
					 // ��ȡÿһ��Ԫ���ֵ
					 for (int j = 0; j < row.getLastCellNum(); j++) {
						 String value = getCellValue(row.getCell(j));
						 if (!value.equals("")) {
							 out(value + " | ",false);
						 }
					 }
					 out("");
				}
			}
		}
		return rowList;
	}

	

	/**
	 * �Զ������ļ���չ�������ö�Ӧ�Ķ�ȡ����
	 * 
	 * @Title: writeExcel
	 * @Date : 2014-9-11 ����01:50:38
	 * @param xlsPath
	 * @throws IOException
	 */
	/*public List<Row> readExcel(String xlsPath) throws IOException{
		
		//��չ��Ϊ��ʱ��
		if (xlsPath.equals("")){
			throw new IOException("�ļ�·������Ϊ�գ�");
		}else{
			File file = new File(xlsPath);
			if(!file.exists()){
				throw new IOException("�ļ������ڣ�");
			}
		}
		
		//��ȡ��չ��
		String ext = xlsPath.substring(xlsPath.lastIndexOf(".")+1);
		
		try {
			
			if("xls".equals(ext)){				//ʹ��xls��ʽ��ȡ
				return readExcel_xls(xlsPath);
			}else if("xls".equals(ext)){		//ʹ��xlsx��ʽ��ȡ
				return readExcel_xlsx(xlsPath);
			}else{									//���γ���xls��xlsx��ʽ��ȡ
				out("��Ҫ�������ļ�û����չ�������ڳ�����xls��ʽ��ȡ...");
				try{
					return readExcel_xls(xlsPath);
				} catch (IOException e1) {
					out("������xls��ʽ��ȡ�����ʧ�ܣ������ڳ�����xlsx��ʽ��ȡ...");
					try{
						return readExcel_xlsx(xlsPath);
					} catch (IOException e2) {
						out("������xls��ʽ��ȡ�����ʧ�ܣ�\n����ȷ�������ļ���Excel�ļ�����������Ȼ�����ԡ�");
						throw e2;
					}
				}
			}
		} catch (IOException e) {
			throw e;
		}
	}*/
	
	/**
	 * �Զ������ļ���չ�������ö�Ӧ��д�뷽��
	 * 
	 * @Title: writeExcel
	 * @Date : 2014-9-11 ����01:50:38
	 * @param rowList
	 * @throws IOException
	 */
	public void writeExcel(List<Row> rowList) throws IOException{
		writeExcel(rowList,excelPath);
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
		String ext = xlsPath.substring(xlsPath.lastIndexOf(".")+1);
		
		try {
			
			if("xls".equals(ext)){				//ʹ��xls��ʽд��
				writeExcel_xls(rowList,xlsPath);
			}else if("xls".equals(ext)){		//ʹ��xlsx��ʽд��
				writeExcel_xlsx(rowList,xlsPath);
			}else{									//���γ���xls��xlsx��ʽд��
				out("��Ҫ�������ļ�û����չ�������ڳ�����xls��ʽд��...");
				try{
					writeExcel_xls(rowList,xlsPath);
				} catch (IOException e1) {
					out("������xls��ʽд�룬���ʧ�ܣ������ڳ�����xlsx��ʽ��ȡ...");
					try{
						writeExcel_xlsx(rowList,xlsPath);
					} catch (IOException e2) {
						out("������xls��ʽд�룬���ʧ�ܣ�\n����ȷ�������ļ���Excel�ļ�����������Ȼ�����ԡ�");
						throw e2;
					}
				}
			}
		} catch (IOException e) {
			throw e;
		}
	}
	
	/**
	 * �޸�Excel��97-03�棬xls��ʽ��
	 * 
	 * @Title: writeExcel_xls
	 * @Date : 2014-9-11 ����01:50:38
	 * @param rowList
	 * @param dist_xlsPath
	 * @throws IOException
	 */
	public void writeExcel_xls(List<Row> rowList, String dist_xlsPath) throws IOException {
		writeExcel_xls(rowList, excelPath,dist_xlsPath);
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
	public void writeExcel_xls(List<Row> rowList, String src_xlsPath, String dist_xlsPath) throws IOException {

		// �ж��ļ�·���Ƿ�Ϊ��
		if (dist_xlsPath == null || dist_xlsPath.equals("")) {
			out("�ļ�·������Ϊ��");
			throw new IOException("�ļ�·������Ϊ��");
		}
		// �ж��ļ�·���Ƿ�Ϊ��
		if (src_xlsPath == null || src_xlsPath.equals("")) {
			out("�ļ�·������Ϊ��");
			throw new IOException("�ļ�·������Ϊ��");
		}

		// �ж��б��Ƿ������ݣ����û�����ݣ��򷵻�
		if (rowList == null || rowList.size() == 0) {
			out("�ĵ�Ϊ��");
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

	/**
	 * �޸�Excel��97-03�棬xls��ʽ��
	 * 
	 * @Title: writeExcel_xls
	 * @Date : 2014-9-11 ����01:50:38
	 * @param rowList
	 * @param dist_xlsPath
	 * @throws IOException
	 */
	public void writeExcel_xlsx(List<Row> rowList, String dist_xlsPath) throws IOException {
		writeExcel_xls(rowList, excelPath , dist_xlsPath);
	}

	/**
	 * �޸�Excel��2007�棬xlsx��ʽ��
	 * 
	 * @Title: writeExcel_xlsx
	 * @Date : 2014-9-11 ����01:50:38
	 * @param rowList
	 * @param xlsPath
	 * @throws IOException
	 */
	public void writeExcel_xlsx(List<Row> rowList, String src_xlsPath, String dist_xlsPath) throws IOException {

		// �ж��ļ�·���Ƿ�Ϊ��
		if (dist_xlsPath == null || dist_xlsPath.equals("")) {
			out("�ļ�·������Ϊ��");
			throw new IOException("�ļ�·������Ϊ��");
		}
		// �ж��ļ�·���Ƿ�Ϊ��
		if (src_xlsPath == null || src_xlsPath.equals("")) {
			out("�ļ�·������Ϊ��");
			throw new IOException("�ļ�·������Ϊ��");
		}

		// �ж��б��Ƿ������ݣ����û�����ݣ��򷵻�
		if (rowList == null || rowList.size() == 0) {
			out("�ĵ�Ϊ��");
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

	/**
	 * //��ȡExcel 2007�棬xlsx��ʽ
	 * 
	 * @Title: readExcel_xlsx
	 * @Date : 2014-9-11 ����11:43:11
	 * @return
	 * @throws IOException
	 */
/*	public List<Row> readExcel_xlsx() throws IOException {
		return readExcel_xlsx(excelPath);
	}*/

	

	/***
	 * ��ȡExcel(97-03�棬xls��ʽ)
	 * 
	 * @throws IOException
	 * 
	 * @Title: readExcel
	 * @Date : 2014-9-11 ����09:53:21
	 */
/*	public List<Row> readExcel_xls() throws IOException {
		return readExcel_xls(excelPath);
	}
*/
	
	/***
	 * ��ȡ��Ԫ���ֵ
	 * 
	 * @Title: getCellValue
	 * @Date : 2014-9-11 ����10:52:07
	 * @param cell
	 * @return
	 */
	private String getCellValue(Cell cell) {
		Object result = "";
		if (cell != null) {
			switch (cell.getCellType()) {
			case Cell.CELL_TYPE_STRING:
				result = cell.getStringCellValue();
				break;
			case Cell.CELL_TYPE_NUMERIC:
				result = cell.getNumericCellValue();
				break;
			case Cell.CELL_TYPE_BOOLEAN:
				result = cell.getBooleanCellValue();
				break;
			case Cell.CELL_TYPE_FORMULA:
				result = cell.getCellFormula();
				break;
			case Cell.CELL_TYPE_ERROR:
				result = cell.getErrorCellValue();
				break;
			case Cell.CELL_TYPE_BLANK:
				break;
			default:
				break;
			}
		}
		return result.toString();
	}

	
	/**
	 * �޸�Excel�������Ϊ
	 * 
	 * @Title: WriteExcel
	 * @Date : 2014-9-11 ����01:33:59
	 * @param wb
	 * @param rowList
	 * @param xlsPath
	 */
	private void writeExcel(Workbook wb, List<Row> rowList, String xlsPath) {

		if (wb == null) {
			out("�����ĵ�����Ϊ�գ�");
			return;
		}

		Sheet sheet = wb.getSheetAt(0);// �޸ĵ�һ��sheet�е�ֵ

		// ���ÿ����д����ô��ӿ�ʼ��ȡ��λ��д���������ȡԴ�ļ����µ��С�
		int lastRowNum = isOverWrite ? startReadRowPos : sheet.getLastRowNum() + 1;
		int t = 0;//��¼������ӵ�����
		out("Ҫ��ӵ�����������Ϊ��"+rowList.size());
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
		out("���м�⵽�ظ�����Ϊ:" + (rowList.size() - t) + " ��׷������Ϊ��"+t);
		
		// ͳһ�趨�ϲ���Ԫ��
		setMergedRegion(sheet);
		
		try {
			// ���½�����д��Excel��
			FileOutputStream outputStream = new FileOutputStream(xlsPath);
			wb.write(outputStream);
			outputStream.flush();
			outputStream.close();
		} catch (Exception e) {
			out("д��Excelʱ�������� ");
			e.printStackTrace();
		}
	}

	/**
	 * ����ĳ�������Ƿ���Excel���д��ڣ�����������
	 * 
	 * @Title: findInExcel
	 * @Date : 2014-9-11 ����02:23:12
	 * @param sheet
	 * @param row
	 * @return
	 */
	private int findInExcel(Sheet sheet, Row row) {
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
	}

	/**
	 * ����һ����Ԫ����ʽ��Ŀ�ĵ�Ԫ����ʽ
	 * 
	 * @param fromStyle
	 * @param toStyle
	 */
	public static void copyCellStyle(CellStyle fromStyle, CellStyle toStyle) {
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

	/**
	 * ��ȡ�ϲ���Ԫ���ֵ
	 * 
	 * @param sheet
	 * @param row
	 * @param column
	 * @return
	 */
	public void setMergedRegion(Sheet sheet) {
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
	

	/**
	 * ��ӡ��Ϣ��
	 * @param msg ��Ϣ����
	 * @param tr ����
	 */
	private void out(String msg){
		if(printMsg){
			out(msg,true);
		}
	}
	/**
	 * ��ӡ��Ϣ��
	 * @param msg ��Ϣ����
	 * @param tr ����
	 */
	private void out(String msg,boolean tr){
		if(printMsg){
			System.out.print(msg+(tr?"\n":""));
		}
	}

	public String getExcelPath() {
		return this.excelPath;
	}

	public void setExcelPath(String excelPath) {
		this.excelPath = excelPath;
	}

	public boolean isNeedCompare() {
		return isNeedCompare;
	}

	public void setNeedCompare(boolean isNeedCompare) {
		this.isNeedCompare = isNeedCompare;
	}

	public int getComparePos() {
		return compareColumnPos;
	}

	public void setComparePos(int comparePos) {
		this.compareColumnPos = comparePos;
	}

	public int getStartReadPos() {
		return startReadRowPos;
	}

	public void setStartReadPos(int startReadPos) {
		this.startReadRowPos = startReadPos;
	}

	public int getEndReadPos() {
		return endReadRowPos;
	}

	public void setEndReadPos(int endReadPos) {
		this.endReadRowPos = endReadPos;
	}

	public boolean isOverWrite() {
		return isOverWrite;
	}

	public void setOverWrite(boolean isOverWrite) {
		this.isOverWrite = isOverWrite;
	}

	public boolean isOnlyReadOneSheet() {
		return onlyReadOneSheet;
	}

	public void setOnlyReadOneSheet(boolean onlyReadOneSheet) {
		this.onlyReadOneSheet = onlyReadOneSheet;
	}

	public int getSelectedSheetIdx() {
		return selectedSheetIndex;
	}

	public void setSelectedSheetIdx(int selectedSheetIdx) {
		this.selectedSheetIndex = selectedSheetIdx;
	}

	public String getSelectedSheetName() {
		return selectedSheetName;
	}

	public void setSelectedSheetName(String selectedSheetName) {
		this.selectedSheetName = selectedSheetName;
	}

	public int getStartSheetIdx() {
		return startSheetIdx;
	}

	public void setStartSheetIdx(int startSheetIdx) {
		this.startSheetIdx = startSheetIdx;
	}

	public int getEndSheetIdx() {
		return endSheetIdx;
	}

	public void setEndSheetIdx(int endSheetIdx) {
		this.endSheetIdx = endSheetIdx;
	}

	public boolean isPrintMsg() {
		return printMsg;
	}

	public void setPrintMsg(boolean printMsg) {
		this.printMsg = printMsg;
	}
	
	private static void saveErrorMessage(String funName,String msg){
		System.out.println(msg);
	}
}