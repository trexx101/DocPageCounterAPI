package com.thetha.ProofByPHD;

import java.io.FileInputStream;

import org.apache.poi.hpsf.PropertySet;
import org.apache.poi.hpsf.SummaryInformation;
import org.apache.poi.poifs.filesystem.DirectoryEntry;
import org.apache.poi.poifs.filesystem.DocumentEntry;
import org.apache.poi.poifs.filesystem.DocumentInputStream;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;
import org.apache.poi.xwpf.extractor.XWPFWordExtractor;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.springframework.boot.SpringApplication;
import org.springframework.boot.autoconfigure.SpringBootApplication;
import org.springframework.boot.context.properties.EnableConfigurationProperties;

import com.thetha.ProofByPHD.model.FileStorageProperties;





@SpringBootApplication
@EnableConfigurationProperties({
	FileStorageProperties.class
})
public class ProofByPhdApplication {

	public static void main(String[] args) {
		SpringApplication.run(ProofByPhdApplication.class, args);
		startApplication();
	}

	private static void startApplication() {
		/*try {
			
            
            XWPFDocument xwpfDocument = new XWPFDocument(new FileInputStream("/home/logia/Desktop/tester.docx"));
			
            System.out.println("*****yo, word count comming next ::> ");
			System.out.println(xwpfDocument.getProperties().getExtendedProperties().getUnderlyingProperties().toString());
		} catch (Exception ex) {
			System.out.println(ex.getMessage());
		}*/
		
	}
}
