package excelUtil;

import java.util.ArrayList;

import org.apache.log4j.Logger;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;

/**
 *@auchor HPC
 *
 */
public class SheetWriteUtil {
	/**
	 * ��־����
	 */
	private static Logger logger = Logger.getLogger("excelLog");
	
	/**
	 * ʵ����sheet
	 */
	private Sheet sheet;
	
	/**
	 * ��sheet������������
	 */
	private ArrayList<Row> allRowList = null;
	
	public SheetWriteUtil(){
		
	}
	

}
