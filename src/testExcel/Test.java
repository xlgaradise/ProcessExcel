package testExcel;

import java.awt.print.Printable;

public class Test {

	public static void main(String[] args) {
		// TODO 自动生成的方法存根

		try {
			boolean b = ExcelUtil.isExcelFile("C:\\Users\\Administrator\\Desktop\\exc");
			if(b)
				System.out.println("done b is true");
			else {
				System.out.println("done b is false");
			}
		} catch (Exception e) {
			// TODO 自动生成的 catch 块
			e.printStackTrace();
		}
	}

}
