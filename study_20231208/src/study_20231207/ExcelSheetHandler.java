package study_20231207;

import java.io.File;
import java.io.InputStream;
import java.util.ArrayList;

import org.apache.poi.ooxml.util.SAXHelper;
import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.ss.util.CellReference;
import org.apache.poi.xssf.eventusermodel.ReadOnlySharedStringsTable;
import org.apache.poi.xssf.eventusermodel.XSSFReader;
import org.apache.poi.xssf.eventusermodel.XSSFSheetXMLHandler;
import org.apache.poi.xssf.eventusermodel.XSSFSheetXMLHandler.SheetContentsHandler;
import org.apache.poi.xssf.model.StylesTable;
import org.apache.poi.xssf.usermodel.XSSFComment;
import org.xml.sax.InputSource;
import org.xml.sax.XMLReader;

public class ExcelSheetHandler implements SheetContentsHandler {

	private int currentCol = -1;
	private int currRowNum = 0;
	private ArrayList<ArrayList<String>> rows = new ArrayList<ArrayList<String>>(); // 실제 엑셀을 파싱해서 담아지는 데이터
	private ArrayList<String> row = new ArrayList<String>();
	private ArrayList<String> header = new ArrayList<String>();

	public static ExcelSheetHandler readExcel(File file) throws Exception {

		ExcelSheetHandler sheetHandler = new ExcelSheetHandler();
		try {

			OPCPackage opc = OPCPackage.open(file);
			XSSFReader xssfReader = new XSSFReader(opc);
			StylesTable styles = xssfReader.getStylesTable();
			ReadOnlySharedStringsTable strings = new ReadOnlySharedStringsTable(opc);

			// 엑셀의 시트를 하나만 가져오기.
			// 여러개일경우 iter문으로 추출해야 함. (iter문으로)
			InputStream inputStream = xssfReader.getSheetsData().next();
			InputSource inputSource = new InputSource(inputStream);
			XSSFSheetXMLHandler handle = new XSSFSheetXMLHandler(styles, strings, sheetHandler, false);

			XMLReader xmlReader = SAXHelper.newXMLReader();
			xmlReader.setContentHandler(handle);

			xmlReader.parse(inputSource);
			inputStream.close();
			opc.close();
		} catch (Exception e) {
			// 에러 발생했을때
		}

		return sheetHandler;

	}

	public ArrayList<ArrayList<String>> getRows() {
		return rows;
	}



	/*
	 * 엑셀 데이터를 처리하기 위한 핸들러 반드시 오버라이딩 해서 구현 해줘야함
	 * 
	 * 3개 메서드 : startRow,cell,endRow
	 * 
	 * stratRow : 하나의 열에 대해 접근을 시작 할때 한번
	 *  cell : 열에 대한 각 셀에 대하여 한번씩
	 *   endRow : 하나의 열에 대해 접근을 끝낼때 한번
	 * 
	 */

	
	/**
	 * 열에 대한 접근을 시작할때 이벤트
	 * 
	 * @param startRow - 접근한 열의 인덱스 (0부터)
	 */
	@Override
	public void startRow(int startRow) {
		this.currentCol = -1;
		this.currRowNum = startRow;  		
	}

	public void hyperlinkCell(String arg0, String arg1, String arg2, String arg3, XSSFComment arg4) {
			//사용 X
	}
	
	
	
	/**
	 * 열에 대한 각 셀에 대한 이벤트(데이터가 있는 셀만 읽음(공배처리 활용))
	 * 
	 * @param columnName - 가로 세로 위치 반환 (EX A1,B2)
	 * @param value      - columnName의 위치에 대한 값 반환
	 */
	@Override
	public void cell(String columnName, String value, XSSFComment var3) {

		int iCol = (new CellReference(columnName)).getCol(); // 각 열에 대한 데이터의 인덱스 1번째 자리 => 0   B1의 인덱스 번호 -> 1 
		int emptyCol = iCol - currentCol - 1;  
		// 공백이 몇개 인지 계산  4번에서 6번으로 옮기면 5번이 공백이니 6-4-1 해서 1이 들어감 

		
		//currentCol부터 iCol까지 공백처리
		for (int i = 0; i < emptyCol; i++) {
			row.add("");
		}	
		
		currentCol = iCol; //현재 셀 위치 저장 하여, 다음 셀까지 공백이 있는지 체크하기 위함
		row.add(value);
	}
	
	
	/**
	 * 열에 대한 접근을 끝낼때 이벤트
	 * 
	 * @param startRow - 접근한 열의 인덱스 (0부터)
	 */

	@Override
	public void endRow(int rowNum) {
		System.out.println("endRow=====>" + rowNum);
		if (rowNum == 0) {
			header = new ArrayList<String>(row);
		} else {
			if (row.size() < header.size()) {
				for (int i = row.size(); i < header.size(); i++) {
					row.add("");
				}
			}
			rows.add(new ArrayList<String>(row));
		}
		row.clear();

	}
	
}
