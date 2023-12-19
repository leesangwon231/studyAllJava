package study_20231207;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.text.NumberFormat;
import java.util.ArrayList;

import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ExcelRead {

	public static void main(String[] args) {
		saxExcelRead();
	}
	
	public static void excelFileRead_1() {
		try {
			// 아래와 같은 방법으로 하는 이유와 방식 차이 찾아보기 **
			// 엑셀파일의 경로와 이름을 통해 POIFSFileSystem을 생성
			// POIFSFileSystem fs = new POIFSFileSystem(new FileInputStream(excelFile)); 
			FileInputStream file = new FileInputStream("C:\\study-20231207\\통합 문서1.xlsx");
			// 엑셀 파일에 대한 워크북 생성
			XSSFWorkbook workbook = new XSSFWorkbook(file);
			
			NumberFormat f = NumberFormat.getInstance();
			f.setGroupingUsed(false);	//지수로 안나오게 설정
			/**
			 *  Double형 데이터가 커지면 데이터에 영문 E와 숫자가 붙어서 표현됨.
			 *  (예: 4.26842654E8) => 지수 표기법
			 *  +) 지수란? 아주 큰 숫자를 표현하기 위해 사용하는 표기법. E 뒤에 붙는 숫자가 10의 승 수를 의미. (예: E8 = 10^8 (10의 8승))
			 *  
			 *  지수 표현법을 제거하는 방법 2가지
			 *  ① NumberFormat의 setGroupingUsed 속성 이용
			 *  	사용법
			 *  		NumberFormat f = NumberFormat.getInstance();
			 *  		f.setGroupingUsed(false);
			 *  
			 *  		f.format(지수인 변수이름);
			 *  
			 *  ② BigDecimal 이용
			 *  	사용법
			 *  		BigDecimal bigDecimal = new BigDecimal(지수인 변수이름);
			 *  		
			 *  		bigDecimal.toString();
			 */
			
			
			// 시트 개수
			int sheetNum = workbook.getNumberOfSheets();
			
			// 시트 수만큼 반복
			for(int s = 0; s < sheetNum; s++) {
				// s번째 시트 읽어오기
				XSSFSheet sheet = workbook.getSheetAt(s);
				
				// 행 개수
				int rows = sheet.getPhysicalNumberOfRows();
				
				// 행 개수만큼 반복
				for(int r = 0 ; r < rows ; r++) {
					// r번째 행 읽어오기
					XSSFRow row = sheet.getRow(r);
					
					// 한 row(행)에 몇 개의 cell(열)이 존재하는지 확인
					int cells = row.getPhysicalNumberOfCells();
					
					System.out.print("|	"+r+"	|");
					
					// 열 개수만큼 반복
					for(int c = 0 ; c < cells; c++) {
						// cell의 순서 번호
						XSSFCell cell = row.getCell(c);
						
						String value = "";
						
						// 빈 cell인지 확인, 수식이 들어간 셀은 걸러지지 않으므로 밑에서 처리
						if(cell!=null) {
							//타입 체크
							switch(cell.getCellType()) {
							case FORMULA:	// 엑셀수식일 경우
								value=cell.getCellFormula() + " : FORMULA";
								break;
							case STRING:	// 문자일 경우 -> " " 공백도 문자열 취급
								value = cell.getStringCellValue() + " : STRING";
								break;
							case NUMERIC:	// 숫자일 경우
								value = f.format(cell.getNumericCellValue()) + " : NUMERIC";
								break;
							case BLANK:		// 빈 셀인 경우
								value = cell.getBooleanCellValue() + " : BLANK";
								break;
							case ERROR:		// 에러일 경우
								value = cell.getErrorCellValue() + " : ERROR";
								break;
							}
						}
						System.out.print("		"+value+"		|");
					}
					System.out.println();
				}
			}
			
			
			/*
			 * 확장성이 없어보여서 주석 처리 : 2023-12-05
			 * 
			 * int rowindex = 0; int columnindex = 0;
			 * 
			 * // 첫번째 시트만 읽어오기 XSSFSheet sheet = workbook.getSheetAt(0);
			 * 
			 * int rows = sheet.getPhysicalNumberOfRows(); for(rowindex = 0; rowindex <
			 * rows; rowindex++ ) { XSSFRow row = sheet.getRow(rowindex);
			 * 
			 * if(row != null) { int cells = row.getPhysicalNumberOfCells(); for(columnindex
			 * = 0; columnindex <= cells; columnindex++) {
			 * 
			 * XSSFCell cell = row.getCell(columnindex); String value = "";
			 * 
			 * if(cell == null) { continue; }else {
			 * 
			 * switch (cell.getCellType()) { case FORMULA: value=cell.getCellFormula();
			 * break; case NUMERIC: // 숫자 형식 value=cell.getNumericCellValue()+""; break;
			 * case STRING: // 문자 형식 value=cell.getStringCellValue()+""; break; case BLANK:
			 * // 엑셀 칸 안의 내용이 비었는지 확인 value=cell.getBooleanCellValue()+""; break; case
			 * ERROR: value=cell.getErrorCellValue()+""; break; } }
			 * System.out.println(rowindex + "번 행 : " + columnindex + "번 열 값은 : " + value);
			 * } } }
			 */
			
			
		} catch (FileNotFoundException e1) {			
			e1.printStackTrace();
		} catch (IOException e2) {
			e2.printStackTrace();
		} catch (IllegalStateException e3) {
			e3.printStackTrace();
		}
	}
	
	public static void saxExcelRead() {

		try {
			String filePath = "C:\\이상원\\농금\\통합\\적재\\test.xlsx";
			File file = new File(filePath);

			ExcelSheetHandler excelSheetHandler = ExcelSheetHandler.readExcel(file);

			// excelDatas >>> [[nero@nate.com, Seoul], [jijeon@gmail.com, Busan], [jy.jeon@naver.com, Jeju]]
			ArrayList<ArrayList<String>> excelDatas = excelSheetHandler.getRows();
			
			System.out.println(excelDatas);
			
			/*
			for(ArrayList<String> dataRow : excelDatas) // row 하나를 읽어온다.
			    for(String str : dataRow){ // cell 하나를 읽어온다.
			        System.out.println(str);
			    }*/
		} catch (Exception e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}

	}

}
