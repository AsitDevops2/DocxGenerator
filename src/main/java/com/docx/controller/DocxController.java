package com.docx.controller;

import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.beans.factory.annotation.Value;
import org.springframework.core.io.Resource;
import org.springframework.http.HttpHeaders;
import org.springframework.http.HttpStatus;
import org.springframework.http.MediaType;
import org.springframework.http.ResponseEntity;
import org.springframework.web.bind.annotation.GetMapping;
import org.springframework.web.bind.annotation.PathVariable;
import org.springframework.web.bind.annotation.PostMapping;
import org.springframework.web.bind.annotation.RequestBody;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RestController;

import com.docx.dto.DocxDto;
import com.docx.dto.Response;
import com.docx.dto.StaticDocxDto;
import com.docx.service.DocxService;

@RestController
@RequestMapping("/docxApi")
public class DocxController {
	
	@Autowired
	private DocxService docxService;
	
	@Value("${download.url}")
	String downloadUrl;
	
	@PostMapping("/generateDocx")
	public Response<String> generateDocx(@RequestBody DocxDto docxDto) {
		String file=docxService.generateDocx(docxDto);
		return new Response<>(HttpStatus.OK.value(), "Your File is ready To Download.Copy the following URL",
				downloadUrl+file);
	}
	
	@GetMapping("/download/{fileName:.+}")
	public ResponseEntity<Object> downloadFile(@PathVariable String fileName) {
		if(!docxService.fileExist(fileName))
			return ResponseEntity.status(HttpStatus.NOT_FOUND.value()).body("File Not Found");
		
		Resource resource=docxService.downloadFile(fileName);
		
		return ResponseEntity.ok()
				.contentType(MediaType.parseMediaType("application/octet-stream"))
				.header(HttpHeaders.CONTENT_DISPOSITION,"attachment;filename=\""+resource.getFilename()+"\"")
				.body(resource);
	}
	
	@PostMapping("/generateAndDownloadDocx")
	public ResponseEntity<Object> generateAndDownloadDocx(@RequestBody DocxDto docxDto) {
		String file=docxService.generateDocx(docxDto);
		Resource resource=docxService.downloadFile(file);
		return ResponseEntity.ok()
				.contentType(MediaType.parseMediaType("application/octet-stream"))
				.header(HttpHeaders.CONTENT_DISPOSITION,"attachment;filename=\""+resource.getFilename()+"\"")
				.body(resource);
	}
	
	@PostMapping("/generateStaticDocx")
	public Response<String> generateStaticDocx(@RequestBody StaticDocxDto docxDto) {
		String file=docxService.generateStaticDocx(docxDto);
		return new Response<>(HttpStatus.OK.value(), "Your File is ready To Download.Copy the following URL",
				downloadUrl+file);
	}
	
}
