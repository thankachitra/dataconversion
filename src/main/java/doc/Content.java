package doc;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.InputStreamReader;
import java.util.Scanner;
import org.apache.poi.xwpf.converter.pdf.PdfConverter;


import org.apache.poi.xwpf.converter.pdf.PdfOptions;
import org.apache.poi.xwpf.usermodel.XWPFDocument;

import org.docx4j.Docx4J;
import org.docx4j.convert.in.Doc;

import org.docx4j.openpackaging.packages.WordprocessingMLPackage;
public class Content {

	public static void main(String a[]) throws Exception {
		try {
			Scanner scanner = new Scanner(new InputStreamReader(System.in));
			System.out.println("Please enter your input docx/doc file name with path: ");
			String inputFile= scanner.nextLine();
			System.out.println("input file name: " + inputFile);
			System.out.println("Please enter your output pdf file name with path: ");
			String outputFile= scanner.nextLine();
			System.out.println("inputFile:" + inputFile + ",outputFile:"+ outputFile);

			FileInputStream in=new FileInputStream(inputFile.toLowerCase());
			File outFile=new File(outputFile);
			outFile.createNewFile();
			FileOutputStream oStream = new FileOutputStream(outFile);

			if (inputFile.endsWith(".docx")){
				/* .docx to .pdf conversion*/
				XWPFDocument document=new XWPFDocument(in);
				PdfOptions options = PdfOptions.create();
				PdfConverter.getInstance().convert(document,oStream,options);
			}
			else if (inputFile.endsWith(".doc")) {
				/* .doc to .pdf conversion*/
				WordprocessingMLPackage wordMLPackage =  Doc.convert(in);
				Docx4J.toPDF(wordMLPackage, oStream);
			}
		}
		catch(Exception e) {
			System.out.println(e);
		}
	}


}


