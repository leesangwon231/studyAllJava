package study_20231211;

import java.io.File;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ExcelWrite {

	public static void main(String[] args) {
		// 엑셀 파일을 쓸 위치 저장
		String filePath = "C:\\study-20231207\\test20231211_2.xlsx";
		String[] names = {"김철수", "홍길동", "이안나", "김수지", "박동희", "손수임", "곽동진", "고길동", "이일표", "조동희"};

		try {
			
			// 엑셀 워크 북 선언
			XSSFWorkbook xWorkbook = null;
			// 엑셀 시트 선언
			XSSFSheet xSheet = null;
			// 엑셀 행 선언
			XSSFRow xRow = null;
			// 엑셀 열 선언
			XSSFCell xCell = null;
			
			
			// 엑셀 워크 북 생성
			/**
			 * 엑셀 파일을 읽거나 쓰기위해 기본으로 필요한 객체
			 * 해당 객체에서 시트를 읽어오고 시트에서 행,열을 찾아 읽거나 쓴다.
			 */
			xWorkbook = new XSSFWorkbook();
			// 시트의 이름을 설정하여 생성
			xSheet = xWorkbook.createSheet("xssSheet_sample");
			
			// 가장 위에 올라올 행 생성 = 헤더가 될 부분
			xRow = xSheet.createRow(0);
			// 헤더에 들어갈 열 생성
			xCell = xRow.createCell(0);
			xCell.setCellValue("번호");
			xCell = xRow.createCell(1);
			xCell.setCellValue("이름");
			xCell = xRow.createCell(2);
			xCell.setCellValue("생일");
			
			
			for(int i = 1; i < 11; i++) {
				int ranNum_y = (int)(Math.random() * 64) + 1960;
				int ranNum_m = (int)(Math.random() * 12) + 1;
				int ranNum_d = (int)(Math.random() * 28) + 1;
				
				String month = ranNum_m < 10 ? "0" + ranNum_m : Integer.toString(ranNum_m);
				String day = ranNum_d < 10 ? "0" + ranNum_d : Integer.toString(ranNum_d);
				
				xRow = xSheet.createRow(i);
				xCell = xRow.createCell(0);
				xCell.setCellValue(i);
				xCell = xRow.createCell(1);
				xCell.setCellValue(names[i-1]);
				xCell = xRow.createCell(2);
				xCell.setCellValue(ranNum_y + "-" + month + "-" + day);
			}
			
			
			File file = new File(filePath);
			FileOutputStream fos;
			
			// FileNotFoundException
			fos = new FileOutputStream(file);
			
			// IOException
			xWorkbook.write(fos);

			if (fos != null)
				fos.close();

		} catch (FileNotFoundException e) {
			e.printStackTrace();
		} catch (IOException e) {
			e.printStackTrace();
		}

	}

}
