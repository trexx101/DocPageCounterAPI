package com.thetha.ProofByPHD.service;

import java.io.*;
import java.math.BigInteger;
import java.net.MalformedURLException;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.nio.file.StandardCopyOption;
import java.util.ArrayList;
import java.util.List;

import com.itextpdf.text.pdf.PdfReader;
import org.apache.poi.hwpf.HWPFDocument;
import org.apache.poi.xwpf.converter.pdf.PdfConverter;
import org.apache.poi.xwpf.converter.pdf.PdfOptions;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.docx4j.XmlUtils;
import org.docx4j.docProps.extended.Properties;
import org.docx4j.jaxb.Context;
import org.docx4j.openpackaging.exceptions.Docx4JException;
import org.docx4j.openpackaging.packages.WordprocessingMLPackage;
import org.docx4j.openpackaging.parts.CustomXmlDataStoragePropertiesPart;
import org.docx4j.openpackaging.parts.DocPropsExtendedPart;
import org.docx4j.openpackaging.parts.Part;
import org.docx4j.openpackaging.parts.WordprocessingML.BibliographyPart;
import org.docx4j.openpackaging.parts.WordprocessingML.MainDocumentPart;
import org.docx4j.openpackaging.parts.WordprocessingML.StyleDefinitionsPart;
import org.docx4j.openpackaging.parts.relationships.Namespaces;
import org.docx4j.openpackaging.parts.relationships.RelationshipsPart;
import org.docx4j.relationships.Relationship;
import org.docx4j.wml.*;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.core.io.Resource;
import org.springframework.core.io.UrlResource;
import org.springframework.stereotype.Service;
import org.springframework.util.StringUtils;
import org.springframework.web.multipart.MultipartFile;

import com.thetha.ProofByPHD.model.DocumentData;
import com.thetha.ProofByPHD.exception.FileStorageException;
import com.thetha.ProofByPHD.model.FileStorageProperties;

import com.thetha.ProofByPHD.exception.MyFileNotFoundException;

@Service
public class DocumentService {
	private final Path fileStorageLocation;

	@Autowired
	public DocumentService(FileStorageProperties fileStorageProperties) {
		this.fileStorageLocation = Paths.get(fileStorageProperties.getUploadDir()).toAbsolutePath().normalize();

		try {
			Files.createDirectories(this.fileStorageLocation);
		} catch (Exception ex) {
			throw new FileStorageException("Could not create the directory where the uploaded files will be stored.",
					ex);
		}
	}

	public DocumentData storeFile(MultipartFile file) {
		DocumentData dd = new DocumentData();
		// Normalize file name
		String fileName = StringUtils.cleanPath(file.getOriginalFilename());

		try {
			// Check if the file's name contains invalid characters
			if (fileName.contains("..")) {
				throw new FileStorageException("Sorry! Filename contains invalid path sequence " + fileName);
			}

			// Copy file to the target location (Replacing existing file with the same name)
			Path targetLocation = this.fileStorageLocation.resolve(fileName);
			Files.copy(file.getInputStream(), targetLocation, StandardCopyOption.REPLACE_EXISTING);
			
			dd.setDocumentName(fileName);
			dd.setDocumentOwner(targetLocation.toString());

			return dd;
		} catch (IOException ex) {
			throw new FileStorageException("Could not store file " + fileName + ". Please try again!", ex);
		}
	}

	public Resource loadFileAsResource(String fileName) {
		try {
			Path filePath = this.fileStorageLocation.resolve(fileName).normalize();
			Resource resource = new UrlResource(filePath.toUri());
			if (resource.exists()) {
				return resource;
			} else {
				throw new MyFileNotFoundException("File not found " + fileName);
			}
		} catch (MalformedURLException ex) {
			throw new MyFileNotFoundException("File not found " + fileName, ex);
		}
	}

	public DocumentData processDocument2(String fileName) {
		File file = new File(fileName);
		File pdfFile = null;
		int noPages = 0;
		DocumentData dd = new DocumentData();

		System.out.println("-------Testing docx flow------");
		try {
			//File file = new File("C:\\Users\\User\\Documents\\Hello.docx");
			WordprocessingMLPackage wordMLPackage = WordprocessingMLPackage.load(file);
			MainDocumentPart mainDocumentPart = wordMLPackage.getMainDocumentPart();
			RelationshipsPart relationshipsPart = wordMLPackage.getMainDocumentPart().getRelationshipsPart();
			List<Relationship> deletions = new ArrayList<Relationship>();

			for (Relationship relationship : relationshipsPart.getRelationships().getRelationship()){
				System.out.println("For Relationship Id=" + relationship.getId()
						+ " Source is " + relationshipsPart.getSourceP().getPartName()
						+ ", Target is " + relationship.getTarget() );

				if (relationship.getTargetMode() != null
						&& relationship.getTargetMode().equals("External") ) {

					continue;
				}

				try {
					Part part = relationshipsPart.getPart(relationship);

					if (part instanceof BibliographyPart) {
						deletions.add(relationship );

						// it is also stored by itemId in a hashmap, so for completeness, delete it from there
						String itemId = getItemId((BibliographyPart)part);
						if (itemId!=null) {
							System.out.println("deleting " + itemId);
							wordMLPackage.getCustomXmlDataStorageParts().remove(itemId);
						}

					}

				} catch (Exception e) {
					throw new Docx4JException("Failed to add parts from relationships", e);
				}
				for ( Relationship r : deletions) {
					System.out.println("deleting " + r.getId() );
					relationshipsPart.removeRelationship(r);
				}
			}

			DocPropsExtendedPart dpep = wordMLPackage.getDocPropsExtendedPart();
			Properties extendedProps = dpep.getContents();
			System.out.println("...Page count --> "+extendedProps.getPages());

			//copy old file
			Document document = mainDocumentPart.getJaxbElement();
			//String str3 = dpep.getXML();
			//Properties extDocPropPart = (Properties) XmlUtils.unmarshalString(dpep.getXML());

			// Create the package and copy style
			WordprocessingMLPackage wordMLPackage2 = WordprocessingMLPackage.createPackage();

			MainDocumentPart newMDP = wordMLPackage2.getMainDocumentPart();
			StyleDefinitionsPart styleDefinitionsPart = newMDP.getStyleDefinitionsPart();
			if (styleDefinitionsPart == null){
				System.out.println("Sheeh!! nothing in that style part");
			}
			//System.out.println(styleDefinitionsPart.getXML());
			String styleXML = styleDefinitionsPart.getXML();
			Styles styles = (Styles) XmlUtils.unmarshalString(styleXML);

			Style style = mainDocumentPart.getStyleDefinitionsPart().getDefaultParagraphStyle();

			Style style1 = setStyleMLA(style, true);
			styles.getStyle().add(style1);
			mainDocumentPart.init();
			System.out.println("produce file..");
			//File docxFile = File.createTempFile("temp-file", ".docx");
			wordMLPackage.save(file);

			try
			{
				System.out.println("trying..");
				// 1) Load docx with POI XWPFDocument

				//XWPFDocument doc = new XWPFDocument(new FileInputStream("C:\\Users\\User\\Documents\\Hello_fixed.docx"));
				XWPFDocument doc = new XWPFDocument(new FileInputStream(file));
				//XWPFDocument document = new XWPFDocument( Data.class.getResourceAsStream( "C:\\Users\\User\\Documents\\Hello_fixed.docx" ) );

				// 2) Convert POI XWPFDocument 2 PDF with iText

				//File pdfFile = new File( "C:\\Users\\User\\Documents\\DocxStructures.pdf" );
				pdfFile = File.createTempFile("temp-file", ".pdf");
				//pdfFile.getParentFile().mkdirs();
				System.out.println("out file created");
				OutputStream out = new FileOutputStream( pdfFile );
				PdfOptions options = PdfOptions.create();
				PdfConverter.getInstance().convert( doc, out, options );
				System.out.println("converted, now closing ...");

				out.close();
			}
			catch ( Throwable e )
			{
				System.out.println("error:: "+e.getMessage());

			}

			PdfReader documents = new PdfReader(new FileInputStream(pdfFile));
			noPages = documents.getNumberOfPages();

			System.out.println("number of pages:: "+noPages);


		} catch (Exception ex) {
			System.out.println(ex.getMessage());
		}
		dd.setDocumentName(file.getName());
		dd.setPageCount(noPages);
		dd.setWordCount(0);

		return dd;
	}

//deprecated
	public DocumentData processDocument(String fileName) {
		File file = new File(fileName);
		DocumentData dd = new DocumentData();
		try {
			if (file.getName().endsWith(".docx")) {

				XWPFDocument doc = new XWPFDocument(new FileInputStream(file.getAbsolutePath()));
				
				dd.setDocumentName("bohoo");
				dd.setWordCount(doc.getProperties().getExtendedProperties().getUnderlyingProperties().getWords());
				dd.setPageCount(doc.getProperties().getExtendedProperties().getUnderlyingProperties().getPages());
				//dd.setWordCount(doc.getProperties().getExtendedProperties().getUnderlyingProperties().toString());
				// textExtractor = new XWPFWordExtractor(doc);
				//System.out.println(doc.getProperties().getExtendedProperties().getUnderlyingProperties().);
			} else if (file.getName().endsWith(".doc")) {
				// textExtractor = new WordExtractor(new FileInputStream(file));
				//HWPFDocument hwpfDocument = new HWPFDocument(new FileInputStream(file.getAbsolutePath()));
				
				//dd.setDocumentName("bohoo");
				//dd.setWordCount(hwpfDocument.getDocProperties().getCWords());
				//System.out.println(hwpfDocument.getDocProperties().getCWords());
			} else {
				throw new IllegalArgumentException("Not a MS Word file.");
			}
			
			
		} catch (IOException e) {
			e.printStackTrace();
		} catch (IllegalArgumentException iae) {
			iae.printStackTrace();
		}

		return dd;
	}

	static Style setStyleMLA(Style style, boolean justify) {
		ObjectFactory factory = Context.getWmlObjectFactory();
		PPr paragraphProperties = factory.createPPr();

		PPrBase.Spacing sp = factory.createPPrBaseSpacing();
		sp.setAfter(BigInteger.ZERO);
		sp.setBefore(BigInteger.ZERO);
		sp.setLine(BigInteger.valueOf(482));
		sp.setLineRule(STLineSpacingRule.AUTO);
		paragraphProperties.setSpacing(sp);

		style.setPPr(paragraphProperties);
		return style;
	}

	public static String getItemId(BibliographyPart entry) {

		String itemId = null;
		if (entry.getRelationshipsPart()==null) {
			return null;
		} else {
			// Look in its rels for rel of @Type customXmlProps (eg @Target="itemProps1.xml")
			Relationship r = entry.getRelationshipsPart().getRelationshipByType(
					Namespaces.CUSTOM_XML_DATA_STORAGE_PROPERTIES);
			if (r==null) {
				System.out.println(".. but that doesn't point to a  customXmlProps part");
				return null;
			}
			CustomXmlDataStoragePropertiesPart customXmlProps =
					(CustomXmlDataStoragePropertiesPart)entry.getRelationshipsPart().getPart(r);
			if (customXmlProps==null) {
				System.out.println(".. but the target seems to be missing?");
				return null;
			} else {
				return customXmlProps.getItemId().toLowerCase();
			}
		}
	}

}
