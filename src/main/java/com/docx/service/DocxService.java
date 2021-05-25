package com.docx.service;

import org.springframework.core.io.Resource;

import com.docx.dto.DocxDto;
import com.docx.dto.StaticDocxDto;

public interface DocxService {
	
	public String generateDocx(DocxDto docxDto);
	public String getFilePath();
	public String getFileName();
	public Resource downloadFile(String fileName);
	public String generateStaticDocx(StaticDocxDto docxDto);
	public boolean fileExist(String fileName);

}
