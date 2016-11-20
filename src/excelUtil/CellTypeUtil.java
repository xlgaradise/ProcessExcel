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
 *��ȡ��Ԫ���ʽ����
 */
public class CellTypeUtil {

	/**
	 *��Ԫ�����ݵĸ�ʽ
	 */
	public static enum TypeEnum{
		STRING,
		NUMERIC,
		/**
		 * ��STRING�����б������
		 */
		DATE_STR,
		/**
		 * ��NUMERIC�����б������
		 */
		DATE_NUM,
		ERROR,
		FORMULA,
		BOOLEAN,
		BLANK,
		/**
		 * δ֪����
		 */
		UNKNOW
	}
	
	/**
	 *��ϸ���ڸ�ʽ
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
		 * yyyy��MM��dd<br>yyyy��MM��dd��<br>yyyy��MM��dd��
		 */
		yyyy_MM_dd_chinese,
		/**
		 * yyyy��MM��
		 */
		yyyy_MM_chinese,
		/**
		 * MM��dd<br>MM��dd��<br>MM��dd��
		 */
		MM_dd_chinese
	}
	
	
	/*--------------�ֶο�ʼ-------------------*/
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
	 * yyyy��MM��dd<br>yyyy��MM��dd��<br>yyyy��MM��dd��
	 */
	public static final String dateRegexString_ymd_chinese = "^[0-9]{4}��[0-9]{1,2}��[0-9]{1,2}[�պ�]?$";
	
	/**
	 * yyyy��MM��
	 */
	public static final String dateRegexString_ym_chinese = "^[0-9]{4}��[0-9]{1,2}��$";

	/**
	 * MM��dd<br>MM��dd��<br>MM��dd��
	 */
	public static final String dateRegexString_md_chinese = "^[0-9]{1,2}��[0-9]{1,2}[�պ�]?$";

	/*--------------�ֶν���-------------------*/



	/**
	 * ��ȡ��Ԫ��ĸ�ʽ
	 * @param cell ��Ԫ��ʵ��
	 * @return TypeEnum���͸�ʽ
	 * @throws IllegalArgumentException ����CellΪnull
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
			if (DateUtil.isCellDateFormatted(cell)) { // ���ڸ�ʽ
				return TypeEnum.DATE_NUM;
			} else if (format == 31 || format == 57 || format == 58 || format == 14) {// �Զ������ڸ�ʽ
				/**
				 * Excel���Զ����������� (yyyy/m/d) format:14 (yyyy��MM��dd��) format:31
				 * (yyyy��MM��) format:57 (MM-dd��MM��dd��) format:58 �����ݲ�ȷ��
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
	 * �б����ڵ�����(�ɰ���/-.�ַ����������պź���)
	 * @param string Ҫƥ�������
	 * @return ���������ͷ���DateEnum���<br>���������������׳��쳣
	 * @throws IllegalArgumentException ���ݲ���DateEnum����
	 */
	public static DateEnum getDateEnum(String string) throws IllegalArgumentException{
		String regex = "[�����պ�]{1}";
		Pattern pattern = Pattern.compile(regex);
		Matcher matcher = pattern.matcher(string);
		if(matcher.find()){ //�������պ���
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
