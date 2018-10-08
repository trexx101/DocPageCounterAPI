package com.thetha.ProofByPHD.model;

public class DocumentData {
	String documentName;
	String documentTitle;
	String documentOwner;
	int wordCount;
	int pageCount;
	
	public DocumentData() {}
	
	public String getDocumentName() {
		return documentName;
	}
	public void setDocumentName(String documentName) {
		this.documentName = documentName;
	}
	public String getDocumentTitle() {
		return documentTitle;
	}
	public void setDocumentTitle(String documentTitle) {
		this.documentTitle = documentTitle;
	}
	public String getDocumentOwner() {
		return documentOwner;
	}
	public void setDocumentOwner(String documentOwner) {
		this.documentOwner = documentOwner;
	}
	public int getWordCount() {
		return wordCount;
	}
	public void setWordCount(int wordCount) {
		this.wordCount = wordCount;
	}

	public int getPageCount() {
		return pageCount;
	}

	public void setPageCount(int pageCount) {
		this.pageCount = pageCount;
	}
	
	
	
	
	

}
