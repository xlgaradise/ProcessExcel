/**
*@auchor HPC
*
*/
package excelUtil;

import java.util.regex.Matcher;
import java.util.regex.Pattern;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.DateUtil;

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
	 * @throws IllegalArgumentException 参数Cell为null
	 */
	@SuppressWarnings("deprecation")
	public static TypeEnum getCellType(Cell cell) throws IllegalArgumentException{
		if(cell == null){
			throw new IllegalArgumentException();
		}

		CellType cellType = cell.getCellTypeEnum();

		switch (cellType) {
		case STRING:
			String value =  cell.getStringCellValue().trim();
			try {
				getDateEnum(value);
				return TypeEnum.DATE_STR;
			} catch (IllegalArgumentException e) {
				return TypeEnum.STRING;
			}
		case NUMERIC:
			short format = cell.getCellStyle().getDataFormat();
			if (DateUtil.isCellDateFormatted(cell)) { // 日期格式
				return TypeEnum.DATE_NUM;
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
	 * @throws IllegalArgumentException 数据不是DateEnum类型
	 */
	public static DateEnum getDateEnum(String string) throws IllegalArgumentException{
		String regex = "[年月日号]{1}";
		Pattern pattern = Pattern.compile(regex);
		Matcher matcher = pattern.matcher(string);
		if(matcher.find()){ //有年月日汉字
			if (Pattern.matches(dateRegexString_ymd_chinese, string)) {
				return DateEnum.yyyy_MM_dd_chinese;
			}else if (Pattern.matches(dateRegexString_ym_chinese, string)) {
				return DateEnum.yyyy_MM_chinese;
			}else if (Pattern.matches(dateRegexString_md_chinese, string)) {
				return DateEnum.MM_dd_chinese;
			}
		}else{
			if(Pattern.matches(dateRegexString_ymd, string)){
				return DateEnum.yyyy_MM_dd;
			}else if (Pattern.matches(dateRegexString_ym, string)) {
				return DateEnum.yyyy_MM;
			}else if (Pattern.matches(dateRegexString_md, string)) {
				return DateEnum.MM_dd;
			}
		}
		throw new IllegalArgumentException();
	}
	
	
}
