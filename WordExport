package com.sgtmcclain.HelloWorldServlet;

import java.io.File;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.math.BigInteger;
import java.util.ArrayList;
import java.util.Calendar;
import java.util.Date;
import java.util.List;

import org.apache.poi.xwpf.model.XWPFHeaderFooterPolicy;
import org.apache.poi.xwpf.usermodel.ParagraphAlignment;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFFooter;
import org.apache.poi.xwpf.usermodel.XWPFHeader;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFRun;
import org.apache.poi.xwpf.usermodel.XWPFTable;
import org.apache.poi.xwpf.usermodel.XWPFTable.XWPFBorderType;
import org.apache.poi.xwpf.usermodel.XWPFTableCell;
import org.apache.poi.xwpf.usermodel.XWPFTableRow;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTPageMar;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTSectPr;


// TODO: Auto-generated Javadoc
/**
 * The Class ViewDetailsWordExport.
 */
public class ViewDetailsWordExport {
	
	/**
	 * The main method.
	 *
	 * @param args the arguments
	 */
	@SuppressWarnings("static-access")
	public static void main(String[] args) {
		
		String docClassification = ("UNCLASSIFIED//FOR OFFICIAL USE ONLY");
		String docAdvisory = ("The user is advised that this International Agreement is not to be used without first obtaining permission from the Office of the Deputy Assistant Secretary of the Army for Defense Exports and Cooperation (DASA (DE&C))");
		
		String agreementIDText = "Agreement: ";
		String agreementID = "ARMY-1496";
		String shortTitle = "PATRIOT/ROLAND FOIA Amendment";
		String objective = "Amendment Number Two to the PATRIOT/ROLAND FOIA is to establish the date-certain for termination of the FOIA as 15 December 2006. The undertakings outlined in the FOIA were completed in December 2005. The Principals agreed that it would take until December 2006 to complete an orderly and comprehensive move to the ,post-FOIA environment.";
		int maxColumns;
		
		
		
		//TODO: Delete when using database
		//Partners Array
		ArrayList<ArrayList<String>> partners = new ArrayList<ArrayList<String>>();
		ArrayList<String> germanyPartner = new ArrayList<String>();
		ArrayList<String> usaPartner = new ArrayList<String>();
		
		germanyPartner.add("Germany");
		germanyPartner.add("$0");
		germanyPartner.add("$0");
		germanyPartner.add("NO");
		germanyPartner.add("YES");
		partners.add(germanyPartner);
		
		usaPartner.add("USA");
		usaPartner.add("$0");
		usaPartner.add("$0");
		usaPartner.add("NO");
		usaPartner.add("NO");
		partners.add(usaPartner);
		
		//POC
		ArrayList<ArrayList<String>> poc = new ArrayList<ArrayList<String>>();
		ArrayList<String> contact = new ArrayList<String>();
		ArrayList<String> contact1 = new ArrayList<String>();
		ArrayList<String> contact2 = new ArrayList<String>();

		
		contact.add("Law Dre");
		contact.add("");
		contact.add("CASA");
		contact.add("703-555-8088");
		contact.add("Law.Dre@gmail.com");
		poc.add(contact);
		
		contact1.add("MR. Man");
		contact1.add("");
		contact1.add("OSD");
		contact1.add("555-555-5555");
		contact1.add("theMan@gmail.com");
		poc.add(contact1);
		
		contact2.add("Mr. Some Guy");
		contact2.add("");
		contact2.add("ORG");
		contact2.add("256-555-0470");
		contact2.add("Some.Guy@gmail.com");
		poc.add(contact2);
		
		
			
			
		
		//Create a blank Document
		XWPFDocument document = new XWPFDocument();
		
		DocumentSettings(document);
		
		//create Header and Footer
		CreateHeaderFooter(document, docClassification, docAdvisory);

	    //AgreementID & Short Title
	    AgreementIDParagraph(document, agreementIDText, agreementID, shortTitle);
	    
	    //Objective
	    ObjectiveParagraph(document, objective);
	    
	    //Agreement Identification (table)

		//create Agreement Identification Table
	    maxColumns = 4;
	    XWPFTable theTable = document.createTable(0, maxColumns);
			
	    	//Set the Borders
	    	XWPFBorderType emptyBorder = null;
			theTable.setInsideHBorder(emptyBorder.NONE, 0, 0, "");
		    theTable.setInsideVBorder(emptyBorder.NONE, 0, 0, "");
	    
		    //Create Header
		    createTableHeader("Agreement Identification", true, theTable, maxColumns);
	
			//create second row
			XWPFTableRow theTableRow = theTable.createRow();
			int tableColumn = 0;
			//create and populate cells for second row
			makeCell(theTableRow, tableColumn++, "Agency:", true, null, true);
			makeCell(theTableRow, tableColumn++, "Parent Agreement ID:", true, null, true);
			makeCell(theTableRow, tableColumn++, "Proponent:", true, null, true);
			makeCell(theTableRow, tableColumn++, "External Agency ID:", true, null, true);
			
			//create third row
			theTableRow = theTable.createRow();
			tableColumn = 0;
			makeCell(theTableRow, tableColumn++, "the agency data", false, null, false);
			//makeCell(theTableRow, tableColumn++, theAgreement.getAgency(), false, null, false);
			makeCell(theTableRow, tableColumn++, "the parent ID data", false, null, false);
			makeCell(theTableRow, tableColumn++, "the proponent data", false, null, false);
			makeCell(theTableRow, tableColumn++, "the external agency ID data", false, null, false);
			
			//create fourth row
			theTableRow = theTable.createRow();
			tableColumn = 0;
			makeCell(theTableRow, tableColumn++, "", false, null, false);
			makeCell(theTableRow, tableColumn++, "", false, null, false);
			makeCell(theTableRow, tableColumn++, "", false, null, false);
			//makeCell(theTableRow, tableColumn++, "Caretaker: ", true, getYesNoText(theAgreement.getCareTaker()), false);
			makeCell(theTableRow, tableColumn++, "Caretaker: ", true, getYesNoText(true), false);
			
		//create Agreement Categories Table
			maxColumns = 4;
			theTable = document.createTable(0,maxColumns);

		
			//Set the Borders
			theTable.setInsideHBorder(emptyBorder.NONE, 0, 0, "");
		    theTable.setInsideVBorder(emptyBorder.NONE, 0, 0, "");
			
			//Create Header
			createTableHeader("Agreement Categories", true, theTable, maxColumns);
			
			//create second row
			theTableRow = theTable.createRow();
			tableColumn = 0;
			
			//create and populate cells for second row
			makeCell(theTableRow, tableColumn++, "Agreement Type:", true, null, true);
			makeCell(theTableRow, tableColumn++, "Agreement Status:", true, null, true);
			makeCell(theTableRow, tableColumn++, "Legal Authority:", true, null, true);
			makeCell(theTableRow, tableColumn++, "Community of Interest", true, null, true);
			
			//create third row
			theTableRow = theTable.createRow();
			tableColumn = 0;
			
			//create and populate cells for third row
			makeCell(theTableRow, tableColumn++, "the Agreement Type", false, null, false);
			makeCell(theTableRow, tableColumn++, "the Parent ID", false, null, false);
			makeCell(theTableRow, tableColumn++, "the Legal Authority", false, null, false);
			makeCell(theTableRow, tableColumn++, "the Community of Interest", false, null, false);
			
		//create Agreement Dates Table
		maxColumns = 4;
		theTable = document.createTable(0, maxColumns);
			
			//Set the Borders
			theTable.setInsideHBorder(emptyBorder.NONE, 0, 0, "");
		    theTable.setInsideVBorder(emptyBorder.NONE, 0, 0, "");
			
			//Create Header
			createTableHeader("Agreement Dates", true, theTable, maxColumns);
			
			//create second row
			theTableRow = theTable.createRow();
			tableColumn = 0;
			
			//create and populate cells for second row
			makeCell(theTableRow, tableColumn++, "Effective Date:", true, null, true);
			makeCell(theTableRow, tableColumn++, "Current Expiration Date:", true, null, true);
			makeCell(theTableRow, tableColumn++, "Original Expiration Date:", true, null, true);
			makeCell(theTableRow, tableColumn++, "Termination Date:", true, null, true);
			
			//create third row
			theTableRow = theTable.createRow();
			tableColumn = 0;
			
			//create and populate cells for third row
			makeCell(theTableRow, tableColumn++, "the Effective Date", false, null, false);
			makeCell(theTableRow, tableColumn++, "the Current Expiration Date", false, null, false);
			makeCell(theTableRow, tableColumn++, "the Original Expiration Date", false, null, false);
			makeCell(theTableRow, tableColumn++, "the Termination Date", false, null, false);
			
			//create fourth row
			theTableRow = theTable.createRow();
			tableColumn = 0;
			
			//create and populate cells for forth row
			makeCell(theTableRow, tableColumn++, "Proposed Date:", true, null, true);
			makeCell(theTableRow, tableColumn++, "Amended Date:", true, null, true);
			makeCell(theTableRow, tableColumn++, "Abandoned  Date:", true, null, true);
			makeCell(theTableRow, tableColumn++, "Does Not Expire:", true, null, true);
			
			//create fifth row
			theTableRow = theTable.createRow();
			tableColumn = 0;
			
			//create and populate cells for fifth row
			makeCell(theTableRow, tableColumn++, "the Proposed Date", false, null, false);
			makeCell(theTableRow, tableColumn++, "the Amended Date", false, null, false);
			makeCell(theTableRow, tableColumn++, "the Abandoned Date", false, null, false);
			makeCell(theTableRow, tableColumn++, getYesNoText(false), false, null, false);
			
		
		document.createParagraph().setSpacingAfter(0);
		
		//create Agreement Partners Table
		maxColumns = 5;
		theTable = document.createTable(0, maxColumns);
		//Create Header
		createTableHeader("Agreement Dates", true, theTable, maxColumns);
		
			//create second row
			theTableRow = theTable.createRow();
			tableColumn = 0;
		
			//create and populate cells for second row
			makeCell(theTableRow, tableColumn++, "Country", true, null, true);
			makeCell(theTableRow, tableColumn++, "Financial Contribution", true, null, true);
			makeCell(theTableRow, tableColumn++, "Non-Financial Contribution", true, null, true);
			makeCell(theTableRow, tableColumn++, "Added", true, null, true);
			makeCell(theTableRow, tableColumn++, "Withdrew", true, null, true);
			
			for(ArrayList<String> partner:partners){
				//create row
				theTableRow = theTable.createRow();
				tableColumn = 0;
				
				//create and populate cells for row
				for(String element:partner){
					makeCell(theTableRow, tableColumn++, element.toString(), false, null, false);
					
				}
			}
		
		document.createParagraph().setSpacingAfter(0);
		//create POC Table
		maxColumns = 5;
		theTable = document.createTable(0, maxColumns);
		//Create Header
		createTableHeader("Points of Contact", true, theTable, maxColumns);
		
			//create second row
			theTableRow = theTable.createRow();
			tableColumn = 0;
		
			//create and populate cells for second row
			makeCell(theTableRow, tableColumn++, "Name", true, null, true);
			makeCell(theTableRow, tableColumn++, "Title", true, null, true);
			makeCell(theTableRow, tableColumn++, "Organization", true, null, true);
			makeCell(theTableRow, tableColumn++, "Phone", true, null, true);
			makeCell(theTableRow, tableColumn++, "Email", true, null, true);
			
			for(ArrayList<String> person:poc){
				//create row
				theTableRow = theTable.createRow();
				tableColumn = 0;
				
				//create and populate cells for row
				for(String element:person){
					makeCell(theTableRow, tableColumn++, element.toString(), false, null, false);
					
				}
			}
		
	    //Write the Document in the file system
	    WriteDocument(document);
		
	}
	
	
	/**
	 * Creates the Table header.
	 *
	 * @param headerText the header text
	 * @param isBold the is bold
	 * @param table the table
	 */
	public static void createTableHeader(String headerText, Boolean isBold, XWPFTable table, int maxColumns){
		//Create grids and column widths
	    table.getCTTbl().addNewTblGrid().addNewGridCol().setW(BigInteger.valueOf(10000));
	    table.getCTTbl().getTblGrid().addNewGridCol().setW(BigInteger.valueOf(10000));
	    
	    //Create the header paragraph and text settings
		@SuppressWarnings("resource")
		XWPFParagraph paragraph = new XWPFDocument().createParagraph();
		paragraph.setAlignment(ParagraphAlignment.CENTER);
		paragraph.setSpacingAfter(0);
		
		XWPFRun run = paragraph.createRun();
		run.setColor("ffffff");
		run.setText(headerText);
		run.setBold(isBold);
		
		//Creating the header row
		XWPFTableRow headerRow = table.getRow(0);
		XWPFTableCell headerCell = table.getRow(0).getCell(0);
		headerRow.getCell(0).setParagraph(paragraph);
		
		headerRow.getCell(0).setVerticalAlignment(XWPFTableCell.XWPFVertAlign.CENTER);
		headerRow.getCell(0).setColor("337ab7");
		
		//Merge the columns
		headerCell.getCTTc().getTcPr().addNewGridSpan().setVal(BigInteger.valueOf((long) maxColumns));
		//Set merged column width
        if (headerCell.getCTTc().getTcPr() == null) {
        	headerCell.getCTTc().addNewTcPr();
        }
        if (headerCell.getCTTc().getTcPr().getTcW()==null) {
        	headerCell.getCTTc().getTcPr().addNewTcW();
        }
		headerCell.getCTTc().getTcPr().getTcW().setW(BigInteger.valueOf(20000L));

	}
	
	/**
	 * Make cell.
	 *
	 * @param tableRow the table row
	 * @param tableColumn the table column
	 * @param cellText the cell text
	 * @param isBold the is bold
	 * @param cellText2 the cell text 2
	 * @param isBold2 the is bold 2
	 */
	private static void makeCell(XWPFTableRow tableRow, int tableColumn, String cellText, Boolean isBold, String cellText2, Boolean isBold2) {
		
		@SuppressWarnings("resource")
		XWPFParagraph paragraph = new XWPFDocument().createParagraph(); 
		XWPFRun run = paragraph.createRun();
		XWPFTableCell cell;
		if(tableRow.getCell(tableColumn) == null){
			cell = tableRow.createCell();
		}else{
			cell = tableRow.getCell(tableColumn);
		}
		
		paragraph.setSpacingAfter(0);
		run.setText(cellText);
		run.setBold(isBold);
		if(cellText2 != null) {
			run = paragraph.createRun();
			run.setText(cellText2);
			run.setBold(isBold2);
		}
		cell.setParagraph(paragraph);
	}
	
	
	/**
	 * Document settings.
	 *
	 * @param document the document
	 */
	private static void DocumentSettings(XWPFDocument document){
		
		//Set the Document Margins
		CTSectPr sectPr = document.getDocument().getBody().addNewSectPr();
	    CTPageMar pageMar = sectPr.addNewPgMar();
	    
	    //Move Header and Footer to the margins
	    pageMar.setHeader(BigInteger.valueOf(370L));
	    pageMar.setFooter(BigInteger.valueOf(370L));
	    
	    //750L corresponds to .5 inches
	    pageMar.setLeft(BigInteger.valueOf(375L));
	    pageMar.setRight(BigInteger.valueOf(375L));
	    pageMar.setTop(BigInteger.valueOf(430L));
	    pageMar.setBottom(BigInteger.valueOf(430L));
		
	}
	
	/**
	 * Creates the header footer.
	 *
	 * @param document the document
	 * @param docClassification the doc classification
	 * @param docAdvisory the doc advisory
	 */
	private static void CreateHeaderFooter(XWPFDocument document, String docClassification, String docAdvisory){
		//Create Header and Footer
	    XWPFHeaderFooterPolicy headerFooterPolicy = document.createHeaderFooterPolicy();
	    XWPFRun run = null;
	    XWPFParagraph paragraph = null;
	    
		//create header start
	    try {
	    	//create header section
			XWPFHeader header = headerFooterPolicy.createHeader(XWPFHeaderFooterPolicy.DEFAULT);
			//Create paragraph section
			paragraph = header.createParagraph();
			paragraph.setAlignment(ParagraphAlignment.CENTER);
			
			//Document Classification
			run = paragraph.createRun();
			run.setText(docClassification);
			run.setBold(true);
			run.addCarriageReturn();
			run.addCarriageReturn();
			//document.removeBodyElement(document.getPosOfParagraph(paragraph));

			//Document Advisory
			run = paragraph.createRun();
			run.setText(docAdvisory);
			run.setItalic(true);
			run.setFontSize(10);
		} catch (IOException e1) {
			// TODO Auto-generated catch block
			e1.printStackTrace();
			//TODO: Change all try catch, e.printStacktrace() to
			//logger.error("Some statement that you want.", e1);
		}
	    
		 //create footer start
	    try {
	    	//create footer section
			XWPFFooter footer = headerFooterPolicy.createFooter(XWPFHeaderFooterPolicy.DEFAULT);
			//Create paragraph section
			paragraph = footer.createParagraph();
			paragraph.setAlignment(ParagraphAlignment.CENTER);
			//Document Classification
			run = paragraph.createRun();
			run.removeCarriageReturn();
			run.setText(docClassification);
			run.setBold(true);
		} catch (IOException e1) {
			// TODO Auto-generated catch block
			e1.printStackTrace();
		}
	}
	
	
	/**
	 * Agreement ID paragraph.
	 *
	 * @param document the document
	 * @param agreementIDText the agreement ID text
	 * @param agreementID the agreement ID
	 * @param shortTitle the short title
	 */
	private static void AgreementIDParagraph(XWPFDocument document, String agreementIDText, String agreementID, String shortTitle){
		XWPFParagraph paragraph = document.createParagraph();
    	XWPFRun run = paragraph.createRun();
    	paragraph.setAlignment(ParagraphAlignment.CENTER);
    	paragraph.setSpacingAfter(0);
    	run.setText(agreementIDText);
    	run.setText(agreementID);
    	run.addCarriageReturn();
    	run.setText(shortTitle);
	}
	
	/**
	 * Objective paragraph.
	 *
	 * @param document the document
	 * @param objective the objective
	 */
	private static void ObjectiveParagraph(XWPFDocument document, String objective){
    	XWPFParagraph paragraph = document.createParagraph();
    	XWPFRun run = paragraph.createRun();
    	
    	run.addCarriageReturn();
    	paragraph.setAlignment(ParagraphAlignment.LEFT);
    	paragraph.setSpacingAfter(0);
    	run.setText(objective);
    	run.setFontSize(10);
	}
	

	
	/**
	 * Write document.
	 *
	 * @param document the document
	 */
	private static void WriteDocument(XWPFDocument document) {
		try (FileOutputStream out = new FileOutputStream (new File("./OtherResources/viewDetailsDocument.docx"));) {
			document.write(out);
			System.out.println("Finshed Creating Document");
		}catch(IOException e) {
			//TODO: Add logger reference.  logger.error("Did not write the view details word document.", e);
		}
	}
	
	
	/**
	 * Gets the yes no text.
	 *
	 * @param isYes the is yes
	 * @return the yes no text
	 */
	private static String getYesNoText(boolean isYes) {
		String yesNo = "";
		if(isYes) {
			yesNo = "YES";
		} else {
			yesNo = "NO";
		}
		return yesNo;
	}
	
}
