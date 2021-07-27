package excel;

import java.io.File;
import java.io.FileInputStream;
import java.net.URL;
import java.util.Iterator;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.Test;

public class TestaExcel {

	@Test
	public void testandoExcel() {
		ClassLoader classloader = org.apache.poi.poifs.filesystem.POIFSFileSystem.class.getClassLoader();
		URL res = classloader.getResource("org/apache/poi/util/POILogger.class");
		String path = res.getPath();
		System.out.println("POI came from " + path);

		FileInputStream planilha = null;
		String planilhaPath = System.getProperty("user.dir").concat("/src/main/resources/TesteExcel.xlsx");
		System.out.println("\n\n" + planilhaPath);

		try {

			File file = new File(planilhaPath);
			planilha = new FileInputStream(file);

			// cria um workbook = planilha toda com todas as abas
			XSSFWorkbook workbook = new XSSFWorkbook(planilha);

			// recupera apenas a primeira aba ou primeira planilha
			XSSFSheet sheet = workbook.getSheetAt(0);

			// restorna todas as linhas da planilha 0 (aba 1)
			Iterator<Row> rowIterator = sheet.iterator();

			// varre todas as linhas da planilha 0
			while (rowIterator.hasNext()) {

				// recebe cada linha da planilha
				Row row = rowIterator.next();

				// pegamos todas as celulas desta linha
				Iterator<Cell> cellIterator = row.iterator();

				// varremos todas as linhas da linha atual
				while (cellIterator.hasNext()) {

					// criamos uma celula ('para cada celula encontrada, fazer algo)
					Cell cell = cellIterator.next();

					switch (cell.getCellType()) {

					case STRING:
						System.out.println("TIPO STRING: " + cell.getStringCellValue());
						break;

					case NUMERIC:
						System.out.println("TIPO NUMBER: " + cell.getNumericCellValue());
						break;

					case FORMULA:
						System.out.println("TIPO FORMULA: " + cell.getCellFormula());
						break;

					default:
						break;
					}
				}
			}

		} catch (Exception e) {
			// TODO: handle exception
			e.getMessage();
		}

	}
}
