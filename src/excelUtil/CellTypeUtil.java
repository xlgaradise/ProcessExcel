/**
*@auchor HPC
*
*/
package excelUtil;

import java.text.ParseException;
import java.text.SimpleDateFormat;
import java.util.Date;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

import org.apache.poi.hssf.usermodel.HSSFDateUtil;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.DateUtil;

import exception.ExcelIllegalArgumentException;
import exception.ExcelNullParameterException;

/**
 *获取单元格格式工具
 */
public class CellTypeUtil {

	/**
	 *单元格内容的格式
	 */
	public static enum TypeEnum{
		STRING,
		NUMERIC,
		/**
		 * 以STRING类型判别的日期
		 */
		DATE_STR,
		/**
		 * 以NUMERIC类型判别的日期
		 */
		DATE_NUM,
		ERROR,
		FORMULA,
		BOOLEAN,
		BLANK,
		/**
		 * 未知类型
		 */
		UNKNOW
	}
	
	/**
	 *详细日期格式
	 */
	public static enum DateEnum{
		/**
		 * yyyy.MM.dd<br>yyyy-MM-dd<br>yyyy/MM/dd
		 */
		yyyy_MM_dd,
		/**
		 * yyyy-MM<br>yyyy/MM
		 */
		yyyy_MM,
		/**
		 * MM-dd<br>MM/dd
		 */
		MM_dd,
		/**
		 * yyyy年MM月dd<br>yyyy年MM月dd日<br>yyyy年MM月dd号
		 */
		yyyy_MM_dd_chinese,
		/**
		 * yyyy年MM月
		 */
		yyyy_MM_chinese,
		/**
		 * MM月dd<br>MM月dd日<br>MM月dd号
		 */
		MM_dd_chinese
	}
	
	
	/*--------------字段开始-------------------*/
	/**
	 * yyyy.MM.dd<br>yyyy-MM-dd<br>yyyy/MM/dd
	 */
	public static final String dateRegexString_ymd = "^[0-9]{4}([/-]|\\.){1}[0-9]{1,2}([/-]|\\.){1}[0-9]{1,2}$";
	
	/**
	 * yyyy-MM<br>yyyy/MM
	 */
	public static final String dateRegexString_ym = "^[0-9]{4}[/-]{1}[0-9]{1,2}$";
	
	/**
	 * MM-dd<br>MM/dd
	 */
	public static final String dateRegexString_md = "^[0-9]{1,2}[/-]{1}[0-9]{1,2}$";
	
	/**
	 * yyyy年MM月dd<br>yyyy年MM月dd日<br>yyyy年MM月dd号
	 */
	public static final String dateRegexString_ymd_chinese = "^[0-9]{4}年[0-9]{1,2}月[0-9]{1,2}[日号]?$";
	
	/**
	 * yyyy年MM月
	 */
	public static final String dateRegexString_ym_chinese = "^[0-9]{4}年[0-9]{1,2}月$";

	/**
	 * MM月dd<br>MM月dd日<br>MM月dd号
	 */
	public static final String dateRegexString_md_chinese = "^[0-9]{1,2}月[0-9]{1,2}[日号]?$";

	/*--------------字段结束-------------------*/



	/**
	 * 获取单元格的格式
	 * @param cell 单元格实例
	 * @return TypeEnum类型格式
	 * @throws ExcelNullParameterException 参数Cell为null
	 */
	@SuppressWarnings("deprecation")
	public static TypeEnum getCellType(Cell cell) throws ExcelNullParameterException{
		if(cell == null){
			throw new ExcelNullParameterException();
		}

		CellType cellType = cell.getCellTypeEnum();

		switch (cellType) {
		case STRING:
			String value =  cell.getStringCellValue().trim();
			try {
				getDateEnum(value);
				return TypeEnum.DATE_STR;
			} catch (ExcelIllegalArgumentException e) {
				return TypeEnum.STRING;
			}
		case NUMERIC:
			short format = cell.getCellStyle().getDataFormat();
			if (DateUtil.isCellDateFormatted(cell)) { // 日期格式
				double d = cell.getNumericCellValue();
				try {
					getFormatDate(d);
					return TypeEnum.DATE_NUM;
				} catch (NullPointerException e) {
					return TypeEnum.NUMERIC;
				}
				
			} else if (format == 31 || format == 57 || format == 58 || format == 14) {// 自定义日期格式
				/**
				 * Excel中自定义日期类型 (yyyy/m/d) format:14 (yyyy年MM月dd日) format:31
				 * (yyyy年MM月) format:57 (MM-dd或MM月dd日) format:58 其他暂不确定
				 */
				return TypeEnum.DATE_NUM;
			} else {
				return TypeEnum.NUMERIC;
			}
		case ERROR:
			return TypeEnum.ERROR;
		case FORMULA:
			return TypeEnum.FORMULA;
		case BOOLEAN:
			return TypeEnum.BOOLEAN;
		case BLANK:
			return TypeEnum.BLANK;
		default:
			return TypeEnum.UNKNOW;
		}
	}
	
	/**
	 * 判别日期的类型(可包含/-.字符或者年月日号汉字)
	 * @param string 要匹配的数据
	 * @return 是日期类型返回DateEnum结果<br>不是日期类型则抛出异常
	 * @throws ExcelIllegalArgumentException 参数不是DateEnum类型
	 */
	public static DateEnum getDateEnum(String string) throws ExcelIllegalArgumentException{
		String regex = "([/-]|\\.){1}";
		Pattern pattern = Pattern.compile(regex);
		Matcher matcher = pattern.matcher(string);
		if(matcher.find()){//包含/.- 
			if (Pattern.matches(dateRegexString_ymd, string)) {
				return DateEnum.yyyy_MM_dd;
			} else if (Pattern.matches(dateRegexString_ym, string)) {
				return DateEnum.yyyy_MM;
			} else if (Pattern.matches(dateRegexString_md, string)) {
				return DateEnum.MM_dd;
			}
		}else{
			String regex2 = "[年月日号]{1}";
			pattern = Pattern.compile(regex2);
			matcher = pattern.matcher(string);
			if(matcher.find()){//有年月日汉字
				if (Pattern.matches(dateRegexString_ymd_chinese, string)) {
					return DateEnum.yyyy_MM_dd_chinese;
				}else if (Pattern.matches(dateRegexString_ym_chinese, string)) {
					return DateEnum.yyyy_MM_chinese;
				}else if (Pattern.matches(dateRegexString_md_chinese, string)) {
					return DateEnum.MM_dd_chinese;
				}
			}
		}
		throw new ExcelIllegalArgumentException();
	}
	
	/**
	 * 获取标准化日期格式
	 * @param dateValue
	 * @return 日期类型格式(yyyy-MM-dd,yyyy-MM,MM-dd)
	 * @throws ExcelIllegalArgumentException 参数不是DateEnum类型
	 */
	public static String getFormatDate(String dateValue) throws ExcelIllegalArgumentException{
		DateEnum dateEnum = getDateEnum(dateValue);
		SimpleDateFormat sdf = null;
		try {
			switch (dateEnum) {
			case yyyy_MM_dd_chinese:
				dateValue = dateValue.replaceAll("[年月]{1}", "-");
				dateValue = dateValue.replaceAll("[日号]?", "");
				if (dateValue.length() != 10) {
					sdf = new SimpleDateFormat("yyyy-MM-dd");
					Date date2 = sdf.parse(dateValue);
					dateValue = sdf.format(date2);
				}
				break;
			case yyyy_MM_chinese:
				dateValue = dateValue.replaceAll("[年月]{1}", "-");
				if (dateValue.length() != 7) {
					sdf = new SimpleDateFormat("yyyy-MM");
					Date date2 = sdf.parse(dateValue);
					dateValue = sdf.format(date2);
				}
				break;
			case MM_dd_chinese:
				dateValue = dateValue.replaceAll("月{1}", "-");
				dateValue = dateValue.replaceAll("[日号]?", "");
				if (dateValue.length() != 5) {
					sdf = new SimpleDateFormat("MM-dd");
					Date date2 = sdf.parse(dateValue);
					dateValue = sdf.format(date2);
				}
				break;
			case yyyy_MM_dd:
				dateValue = dateValue.replaceAll("([/-]|\\.){1}", "-");
				if (dateValue.length() != 10) {
					sdf = new SimpleDateFormat("yyyy-MM-dd");
					Date date2 = sdf.parse(dateValue);
					dateValue = sdf.format(date2);
				}
				break;
			case yyyy_MM:
				dateValue = dateValue.replaceAll("[/-]{1}", "-");
				if (dateValue.length() != 7) {
					sdf = new SimpleDateFormat("yyyy-MM-dd");
					Date date2 = sdf.parse(dateValue);
					dateValue = sdf.format(date2);
				}
				break;
			case MM_dd:
				dateValue = dateValue.replaceAll("[/-]{1}", "-");
				if (dateValue.length() != 5) {
					sdf = new SimpleDateFormat("yyyy-MM-dd");
					Date date2 = sdf.parse(dateValue);
					dateValue = sdf.format(date2);
				}
				break;
			default:
				break;
			}
		} catch (ParseException e) {
			e.printStackTrace();
		}
		return dateValue;
	}
	
	/**
	 * 获取标准化日期格式
	 * @param dateValue
	 * @return 日期类型格式(yyyy-MM-dd)
	 * @throws NullPointerException 日期解析错误
	 */
	public static String getFormatDate(double dateValue) throws NullPointerException{
		Date date = DateUtil.getJavaDate(dateValue);
		if(date == null){
			date = HSSFDateUtil.getJavaDate(dateValue);
		}
		String string = new SimpleDateFormat("yyyy-MM-dd").format(date);
		return string;
	}
}
