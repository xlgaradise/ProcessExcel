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
	 * 日志工具
	 */
	private static Logger logger = Logger.getLogger("excelLog");
	
	/**
	 * 实例的sheet
	 */
	private Sheet sheet;
	
	/**
	 * 该sheet的所有行数据
	 */
	private ArrayList<Row> allRowList = null;
	
	public SheetWriteUtil(){
		
	}
	

}
