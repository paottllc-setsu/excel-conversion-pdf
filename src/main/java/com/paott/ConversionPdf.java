package com.paott;

import java.io.FileInputStream;
import java.io.IOException;
import java.io.InputStream;

import org.apache.pdfbox.pdmodel.PDDocument;
import org.apache.pdfbox.pdmodel.PDPage;
import org.apache.pdfbox.pdmodel.PDPageContentStream;
import org.apache.pdfbox.pdmodel.common.PDRectangle;
import org.apache.pdfbox.pdmodel.font.PDType0Font;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ConversionPdf {

	public static void main(String[] args) {
		if (args.length != 2) {
			System.err.println("Usage: java -jar excel-conversion-pdf.jar <input.xlsx> <output.pdf>");
			System.exit(0);
		}

		String inputExcel = args[0];
		String outputPdf = args[1];

		try {
			convertExcelToPdf(inputExcel, outputPdf);
			System.out.println("PDF化に成功しました。");
			System.exit(1);
		} catch (IOException e) {
			System.err.println("PDF化に失敗しました : " + e.getMessage());
			e.printStackTrace();
			System.exit(2);
		}
	}

	public static void convertExcelToPdf(String inputExcel, String outputPdf) throws IOException {
		try (InputStream inp = new FileInputStream(inputExcel);
				Workbook workbook = new XSSFWorkbook(inp);
				PDDocument document = new PDDocument()) {

			// すべてのシートをループ処理
			for (int sheetIndex = 0; sheetIndex < workbook.getNumberOfSheets(); sheetIndex++) {
				Sheet sheet = workbook.getSheetAt(0); // 最初のシートを取得
				// 横向きA4サイズのPDRectangleを作成
				PDRectangle landscapeA4 = new PDRectangle(PDRectangle.A4.getHeight(), PDRectangle.A4.getWidth());
				PDPage page = new PDPage(landscapeA4); // 横向きA4サイズのPDFページを作成
				document.addPage(page);

				try (PDPageContentStream contentStream = new PDPageContentStream(document, page)) {
					PDType0Font font = PDType0Font.load(document, ConversionPdf.class.getResourceAsStream("/NotoSansJP-Regular.ttf")); // MSゴシックフォントを指定

					//int rowNum = 700;
					// 横向きA4サイズの座標系に合わせて座標を調整
					int rowNum = (int) landscapeA4.getWidth() - 300; // Y座標を調整
					for (Row row : sheet) {
						int colNum = 10;
						for (Cell cell : row) {
							String cellValue = cell.toString();
							CellStyle cellStyle = cell.getCellStyle(); // Excelのフォントサイズを取得
							contentStream.beginText();
							Font excelFont = workbook.getFontAt(cellStyle.getFontIndex());
							short fontSize = excelFont.getFontHeightInPoints();
							contentStream.setFont(font, fontSize); // PDFにフォントサイズを適用
							contentStream.newLineAtOffset(colNum, rowNum);
							contentStream.showText(cellValue);
							contentStream.endText();
							colNum += 100;
						}
						rowNum -= 20; // Y座標の減少量を調整
					}
				} catch (IOException e) {
					System.err.println("フォントのローディングに失敗しました：" + e.getMessage());
					System.exit(3);
				}
				document.save(outputPdf);
			}
		}
	}
}