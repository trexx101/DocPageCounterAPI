package com.thetha.ProofByPHD.service;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.net.MalformedURLException;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.nio.file.StandardCopyOption;

import org.apache.poi.hwpf.HWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
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

	public DocumentData processDocument(String fileName) {
		File file = new File(fileName);
		DocumentData dd = new DocumentData();
		try {
			if (file.getName().endsWith(".docx")) {

				XWPFDocument doc = new XWPFDocument(new FileInputStream(file.getAbsolutePath()));
				
				dd.setDocumentName("bohoo");
				dd.setWordCount(doc.getProperties().getExtendedProperties().getWords());
				// textExtractor = new XWPFWordExtractor(doc);
			} else if (file.getName().endsWith(".doc")) {
				// textExtractor = new WordExtractor(new FileInputStream(file));
				HWPFDocument hwpfDocument = new HWPFDocument(new FileInputStream(file.getAbsolutePath()));
				
				dd.setDocumentName("bohoo");
				dd.setWordCount(hwpfDocument.getDocProperties().getCWords());
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

}
