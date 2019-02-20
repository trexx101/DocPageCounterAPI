package com.thetha.ProofByPHD;

import java.io.*;
import java.math.BigInteger;
import java.util.*;

import com.itextpdf.text.pdf.PdfReader;
import org.apache.poi.xwpf.converter.pdf.PdfConverter;
import org.apache.poi.xwpf.converter.pdf.PdfOptions;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.docx4j.Docx4J;
import org.docx4j.XmlUtils;
import org.docx4j.convert.out.ConversionFeatures;
import org.docx4j.convert.out.FOSettings;
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


import org.docx4j.toc.TocException;
import org.docx4j.toc.TocPageNumbersHandler;
import org.docx4j.wml.*;
import org.springframework.boot.SpringApplication;
import org.springframework.boot.autoconfigure.SpringBootApplication;
import org.springframework.boot.context.properties.EnableConfigurationProperties;
import com.thetha.ProofByPHD.model.FileStorageProperties;


import javax.xml.bind.JAXBContext;
import javax.xml.parsers.SAXParser;
import javax.xml.parsers.SAXParserFactory;


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

		System.out.println("-------Testing docx flow------");
		/*try {
			File file = new File("C:\\Users\\User\\Documents\\Hello.docx");
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

            //System.out.println("...."+mainDocumentPart.getStyleDefinitionsPart().getXML());
            wordMLPackage.save(new java.io.File("C:\\Users\\User\\Documents\\Hello_fixed.docx") );


            try
            {
                System.out.println("trying..");
                // 1) Load docx with POI XWPFDocument
                XWPFDocument doc = new XWPFDocument(new FileInputStream("C:\\Users\\User\\Documents\\Hello_fixed.docx"));
                //XWPFDocument document = new XWPFDocument( Data.class.getResourceAsStream( "C:\\Users\\User\\Documents\\Hello_fixed.docx" ) );

                // 2) Convert POI XWPFDocument 2 PDF with iText
                File outFile = new File( "C:\\Users\\User\\Documents\\DocxStructures.pdf" );
                //outFile.getParentFile().mkdirs();
                System.out.println("out file created");
                OutputStream out = new FileOutputStream( outFile );
                PdfOptions options = PdfOptions.create();
                PdfConverter.getInstance().convert( doc, out, options );
                System.out.println("converted, now closing ...");

                out.close();
            }
            catch ( Throwable e )
            {
                System.out.println("error:: "+e.getMessage());
                //e.printStackTrace();
            }

            PdfReader documents = new PdfReader(new FileInputStream(new File("C:\\Users\\User\\Documents\\DocxStructures.pdf")));
            int noPages = documents.getNumberOfPages();

            System.out.println("number of pages:: "+noPages);


		} catch (Exception ex) {
			System.out.println(ex.getMessage());
		}*/
		
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

    private static MainDocumentPart formatDoc(WordprocessingMLPackage wordMLPackage) {
        MainDocumentPart mdp = wordMLPackage.getMainDocumentPart();

        for ( org.docx4j.wml.Style s : mdp.getStyleDefinitionsPart().getJaxbElement().getStyle() ) {
            if (s.isDefault()) {
                System.out.println("Style wit h name " + s.getName().getVal() + ", id '" + s.getStyleId() + "' is default " + s.getType() + " style");

            }
        }

        mdp.getStyleDefinitionsPart().getJaxbElement().getStyle().stream().forEach((Style s) -> {
            //System.out.println("style typeee: "+s.getType());

            switch (s.getStyleId()) {
                case "Normal":
                    setStyleMLA(s, true);
                    break;
                case "Emphasis":
                    setStyleMLA(s, false);
                    break;
            }
        });
        return mdp;
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

    private static Map<String, Integer> getPageNumbersMapViaFOP(WordprocessingMLPackage wordMLPackage) throws TocException {

        System.out.println("getPageNumbersMapViaFOP() starting..");

        long start = System.currentTimeMillis();

        FOSettings foSettings = Docx4J.createFOSettings();
        foSettings.setWmlPackage(wordMLPackage);
        String MIME_FOP_AREA_TREE   = "application/X-fop-areatree"; // org.apache.fop.apps.MimeConstants
        foSettings.setApacheFopMime(MIME_FOP_AREA_TREE);

        foSettings.getFeatures().add(ConversionFeatures.PP_PDF_APACHEFOP_DISABLE_PAGEBREAK_LIST_ITEM); // in 3.0.1, this is off by default

        /*if (log.isDebugEnabled()) {
            foSettings.setFoDumpFile(new java.io.File(System.getProperty("user.dir") + "/Toc.fo"));
        }*/

        ByteArrayOutputStream os = new ByteArrayOutputStream();
         boolean foViaXSLT = false;
        try {
            if (foViaXSLT) {
                Docx4J.toFO(foSettings, os, Docx4J.FLAG_EXPORT_PREFER_XSL);
            } else {
                Docx4J.toFO(foSettings, os, Docx4J.FLAG_EXPORT_PREFER_NONXSL);
            }

            long end = System.currentTimeMillis();
            float timing = (end-start)/1000;
            if (foViaXSLT) {
                System.out.println("Time taken (AT via XSLT): " + Math.round(timing) + " sec");
            } else {
                System.out.println("Time taken (AT via non XSLT): " + Math.round(timing) + " sec");
            }

//            start = System.currentTimeMillis();
            InputStream is = new ByteArrayInputStream(os.toByteArray());
            SAXParserFactory factory = SAXParserFactory.newInstance();
            SAXParser saxParser = factory.newSAXParser();
            TocPageNumbersHandler tpnh = new TocPageNumbersHandler();
            saxParser.parse(is, tpnh);

            // Negligible
//            end = System.currentTimeMillis();
//            timing = (end-start)/1000;
//            log.debug("Time taken (parse step): " + Math.round(timing) + " sec");

            return tpnh.getPageNumbers();

        } catch (Exception e) {
            throw new TocException(e.getMessage(),e);
        }

    }

    /**
	 * Converts recources inside of the docx to enteries in fidus
	 *
	 * @param wordMLPackage
	 *
	private void setBibliographyResources(WordprocessingMLPackage wordMLPackage) {
		HashMap<String, org.docx4j.openpackaging.parts.CustomXmlPart> mp = wordMLPackage.getCustomXmlDataStorageParts();
		Iterator<Map.Entry<String, org.docx4j.openpackaging.parts.CustomXmlPart>> it = mp.entrySet().iterator();
		while (it.hasNext()) {
			Map.Entry<String, org.docx4j.openpackaging.parts.CustomXmlPart> pair = (Map.Entry<String, org.docx4j.openpackaging.parts.CustomXmlPart>) it
					.next();
			if (pair.getValue() instanceof org.docx4j.openpackaging.parts.WordprocessingML.BibliographyPart) {
				org.docx4j.openpackaging.parts.WordprocessingML.BibliographyPart bibliographyPart = (org.docx4j.openpackaging.parts.WordprocessingML.BibliographyPart) pair
						.getValue();
				try {
					JAXBElement<org.docx4j.bibliography.CTSources> bibliography = bibliographyPart.getContents();
					this.bibliographyResources = bibliography.getValue().getSource();
				} catch (Docx4JException e) {
					System.err.println("\nProblem in reading bibliography part.");
					e.printStackTrace();
				}
				break;
			}
		}

	}*/
}
