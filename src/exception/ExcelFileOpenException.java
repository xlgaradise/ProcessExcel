
package exception;

/**
*@auchor HPC
*@encoding GBK
*/

/**
 *�ļ��򿪴���
 */
public class ExcelFileOpenException extends ExcelBaseException {
	
	private Exception exception;

	/**
	 * 
	 */
	private static final long serialVersionUID = 1L;

	public ExcelFileOpenException(Exception exception){
		this.exception = exception;
	}
	
	public Exception getException(){
		return exception;
	}
}
