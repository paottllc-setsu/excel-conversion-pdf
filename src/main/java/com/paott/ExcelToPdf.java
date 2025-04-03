package com.paott;

import java.io.File;

import org.jodconverter.core.DocumentConverter;
import org.jodconverter.core.office.OfficeManager;
import org.jodconverter.local.LocalConverter;
import org.jodconverter.local.office.LocalOfficeManager;

public class ExcelToPdf {

	public static void main(String[] args) {
		if (args.length != 2) {
			System.err.println("Usage: java -jar excel-conversion-pdf.jar <input.xlsx> <output.pdf>");
			System.exit(1);
		}

		String inputExcel = args[0];
		String outputPdf = args[1];

		try {
			convertExcelToPdf(inputExcel, outputPdf);
			System.out.println("PDF化に成功しました。");
			System.exit(0);
		} catch (Exception e) {
			System.err.println("PDF化に失敗しました: " + e.getMessage());
			e.printStackTrace();
			System.exit(2);
		}
	}

	public static void convertExcelToPdf(String inputExcel, String outputPdf) throws Exception {
		OfficeManager officeManager = null;
		//officeManager = LocalOfficeManager.builder().officeHome("C:/Program Files/LibreOffice").build();
		//officeManager = LocalOfficeManager.builder().install().build();
		//officeManager = LocalOfficeManager.make();
		officeManager = LocalOfficeManager.builder()
				.officeHome("C:/Program Files/LibreOffice")
				.portNumbers(2003) // ポートの競合が起こる場合は変更する
				.build();
		officeManager.start();

		DocumentConverter converter = LocalConverter.builder().officeManager(officeManager).build();
		converter.convert(new File(inputExcel)).to(new File(outputPdf)).execute();

		officeManager.stop();
	}
}