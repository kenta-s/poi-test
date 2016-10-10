import java.io.*;
import java.util.*;
import org.apache.poi.hssf.usermodel.*;
import org.apache.poi.hssf.util.HSSFColor;
 
public class Hello {
	public static void main(String[] arg){
		//ワークブックの生成
		HSSFWorkbook workbook = new HSSFWorkbook();
		//ワークシートの生成
		HSSFSheet sheet = workbook.createSheet("HelloWorld");
		//Rowの生成
		HSSFRow row = sheet.createRow(0);
		//cellの生成
		@SuppressWarnings("deprecation")
		HSSFCell cell = row.createCell((short)0);
 
		//cellスタイルの生成
		HSSFCellStyle st = workbook.createCellStyle();
 
		//フォントの生成
		HSSFFont fnt = workbook.createFont();
		fnt.setFontName("MS 明朝");
		fnt.setFontHeightInPoints((short)48);
		fnt.setColor((short)HSSFColor.AQUA.index);
 
		//cellスタイルにフォント設定
		st.setFont(fnt);
 
		//cellにスタイル設定
		cell.setCellStyle(st);
 
		//cellに値設定
		cell.setCellValue("Hello World♪");
 
		//ワークブック書き出し
		FileOutputStream out = null;
 
		try {
			out = new FileOutputStream(
					"HelloWorld_Book1.xls");
			workbook.write(out);
		} catch (IOException e) {
			System.out.println(e.toString());
		} finally {
			try {
				out.close();
			} catch (IOException e) {
				System.out.println(e.toString());
			}
		}
	}
 
}

