package com.docx.service.serviceImpl;

import java.io.File;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.OutputStream;
import java.net.MalformedURLException;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.util.HashMap;
import java.util.Map;
import java.util.SortedMap;
import java.util.TreeMap;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

import org.apache.commons.io.FilenameUtils;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.xwpf.usermodel.ParagraphAlignment;
import org.apache.poi.xwpf.usermodel.UnderlinePatterns;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFRun;
import org.apache.poi.xwpf.usermodel.XWPFTable;
import org.apache.poi.xwpf.usermodel.XWPFTableCell;
import org.apache.poi.xwpf.usermodel.XWPFTableRow;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.springframework.beans.factory.annotation.Value;
import org.springframework.core.io.Resource;
import org.springframework.core.io.UrlResource;
import org.springframework.stereotype.Service;

import com.docx.dto.DocxDto;
import com.docx.dto.StaticDocxDto;
import com.docx.service.DocxService;
import com.fasterxml.jackson.databind.ObjectMapper;

@Service
public class DocxServiceImpl implements DocxService{
	
	Logger logger=LoggerFactory.getLogger(DocxServiceImpl.class);
	
	private File myFile=null;
	private FileOutputStream out=null;
	
	@Value("${directory.generated}")
	String generatedDir;
	
	public String getFilePath() {
		return FilenameUtils.getFullPathNoEndSeparator(this.myFile.getAbsolutePath());
	}
	
	public String getFileName() {
		return this.myFile.getName();
	}
	
	public Map<String, String> convertObjTomap(DocxDto docxDto) {
		ObjectMapper mapper=new ObjectMapper();
		@SuppressWarnings("unchecked")
		Map<String, String> data=mapper.convertValue(docxDto, HashMap.class);
		
		return data;
	}
	
	
	@Override
	public String generateDocx(DocxDto docxDto) {
		
		//Creating directory named generated if not already created
		if (!Paths.get(generatedDir).toFile().exists()) {
			try {
				Files.createDirectories(Paths.get(generatedDir));
			} catch (IOException e) {
				logger.error(e.getMessage());
				e.printStackTrace();
			}
		}
		
		this.myFile=new File(generatedDir + "/createdWord.docx");
		try {
			this.out = new FileOutputStream(this.myFile);
		} catch (FileNotFoundException e) {
			logger.error(e.getMessage());
			e.printStackTrace();
		}
				
		Map<String, String> data=convertObjTomap(docxDto);
		
		//getting template.docx from template directory
		try {
			replace("./template/template.docx", data, this.out);
			
			this.out.close();
		} catch (Exception e) {
			logger.error(e.getMessage());
			e.printStackTrace();
		}
		
		return getFileName();
	}
	
	@SuppressWarnings("resource")
	private void replace(String inFile, Map<String, String> data, OutputStream out) {
	    XWPFDocument doc=new XWPFDocument();
		try {
			doc = new XWPFDocument(OPCPackage.open(inFile));
		} catch (InvalidFormatException|IOException e) {
			logger.error(e.getMessage());
			e.printStackTrace();
		}
	    
	    for (XWPFParagraph p : doc.getParagraphs()) {
	        replace2(p, data);
	    }
	    
	    for (XWPFTable tbl : doc.getTables()) {
	        for (XWPFTableRow row : tbl.getRows()) {
	            for (XWPFTableCell cell : row.getTableCells()) {
	                for (XWPFParagraph p : cell.getParagraphs()) {
	                    replace2(p, data);
	                }
	            }
	        }
	    }
	    
	    try {
			doc.write(out);
		} catch (IOException e) {
			logger.error(e.getMessage());
			e.printStackTrace();
		}
	}

	private void replace2(XWPFParagraph para, Map<String, String> data) {
	    String paraText = para.getText(); // complete paragraph as string
	    
	    // if paragraph does not include our pattern, ignore
	    if (paraText.contains("${")) 
	    { 
	        TreeMap<Integer, XWPFRun> posRuns = getPosToRuns(para);
	        Pattern pattern = Pattern.compile("\\$\\{(.+?)\\}");
	        Matcher matcher = pattern.matcher(paraText);
	        
	        // for all patterns in the paragraph
	        while (matcher.find()) 
	        { 
	            String group = matcher.group(1);  // extract key start and end pos
	            int start = matcher.start(1);
	            int end = matcher.end(1);
	            String key = group;
	            String value = data.get(key);
	            if (value == null)
	            	value = "";
	            
	            SortedMap<Integer, XWPFRun> range = posRuns.subMap(start - 2, true, end + 1, true); // get runs which contain the pattern
	            // found $
	            boolean foundDollerSymbol = false;
	            // found {
	            boolean foundOpeningCurly = false; 
	            // found }
	            boolean foundClosingCurly = false; 
	            XWPFRun prevRun = null; // previous run handled in the loop
	            XWPFRun foundOpeningCurlyRun = null; // run in which { was found
	            int foundOpeningCurlyPos = -1; // pos of { within above run
	            
	            for (XWPFRun xwpfRun : range.values())
	            {
	                if (xwpfRun == prevRun)
	                    continue; // this run has already been handled
	                if (foundClosingCurly)
	                    break; // done working on current key pattern
	                
	                prevRun = xwpfRun;
	                
	                for (int k = 0;; k++) { // iterate over texts of run xwpfRun
	                    if (foundClosingCurly)
	                        break;
	                    String txt = null;
	                    try {
	                        txt = xwpfRun.getText(k); // note: should return null, but throws exception if the text does not exist
	                    } catch (Exception ex) {
	                    	logger.error(ex.getMessage());
	                    	ex.printStackTrace();
	                    }
	                    if (txt == null)
	                        break; // no more texts in the run, exit loop
	                    if (txt.contains("$") && !foundDollerSymbol) {  // found $, replace it with value from data map
	                        txt = txt.replaceFirst("\\$", value);
	                        foundDollerSymbol = true;
	                    }
	                    if (txt.contains("{") && !foundOpeningCurly && foundDollerSymbol) {
	                        foundOpeningCurlyRun = xwpfRun; // found { replace it with empty string and remember location
	                        foundOpeningCurlyPos = txt.indexOf('{');
	                        txt = txt.replaceFirst("\\{", "");
	                        foundOpeningCurly = true;
	                    }
	                    if (foundDollerSymbol && foundOpeningCurly && !foundClosingCurly) { // find } and set all chars between { and } to blank
	                        if (txt.contains("}"))
	                        {
	                            if (xwpfRun == foundOpeningCurlyRun) // complete pattern was within a single run
	                                txt = txt.substring(0, foundOpeningCurlyPos)+txt.substring(txt.indexOf('}'));
	                            else // pattern spread across multiple runs
	                                txt = txt.substring(txt.indexOf('}'));
	                        }
	                        else if (xwpfRun == foundOpeningCurlyRun) // same run as { but no }, remove all text starting at {
	                            txt = txt.substring(0,  foundOpeningCurlyPos);
	                        else
	                            txt = ""; // run between { and }, set text to blank
	                    }
	                    if (txt.contains("}") && !foundClosingCurly) {
	                        txt = txt.replaceFirst("\\}", "");
	                        foundClosingCurly = true;
	                    }
	                    xwpfRun.setText(txt, k);
	                }
	            }
	        }

	    }
	}

	private TreeMap<Integer, XWPFRun> getPosToRuns(XWPFParagraph paragraph) {
	    int pos = 0;
	    TreeMap<Integer, XWPFRun> map = new TreeMap<Integer, XWPFRun>();
	    for (XWPFRun run : paragraph.getRuns()) {
	        String runText = run.text();
	        if (runText != null && runText.length() > 0) {
	            for (int i = 0; i < runText.length(); i++) {
	                map.put(pos + i, run);
	            }
	            pos += runText.length();
	        }

	    }
	    return map;
	}
	
	public String generateStaticDocx(StaticDocxDto docxDto) {
		//Creating directory named generated if not already created
		if (!Paths.get(generatedDir).toFile().exists()) {
			try {
				Files.createDirectories(Paths.get(generatedDir));
			} catch (IOException e) {
				logger.error(e.getMessage());
				e.printStackTrace();
			}
		}
				
		this.myFile=new File("./generated/" + "StaticWord.docx");
		
		try {
			this.out = new FileOutputStream(this.myFile);
		} catch (FileNotFoundException e) {
			logger.error(e.getMessage());
			e.printStackTrace();
		}
		
		XWPFDocument document=null;
		try {
			document = new XWPFDocument();
			
			//setting paragraph text style
			XWPFParagraph heading = document.createParagraph();
			heading.setAlignment(ParagraphAlignment.CENTER);
			XWPFRun headingR = heading.createRun();
			headingR.setText(docxDto.getTitle());
			headingR.setFontFamily("heading 1");
			headingR.setFontSize(16);
			
			
			 //create Paragraph    
	        XWPFParagraph firstPara = document.createParagraph();
	        XWPFRun run = firstPara.createRun();
	        run.setText(docxDto.getFirstPara());
	        run.setFontSize(14);
	        run.setFontFamily("times new roman");
	        
	        
	        XWPFParagraph secondPara = document.createParagraph();
	        XWPFRun run2 = secondPara.createRun();
	        run2.setText(docxDto.getSecondPara());
	        run2.setItalic(true);
	        run2.setColor("F80636");
	        run2.setBold(true);
	        run2.setUnderline(UnderlinePatterns.SINGLE);
	        
	        XWPFParagraph thirdPara = document.createParagraph();
	        thirdPara.setAlignment(ParagraphAlignment.LEFT);
	        XWPFRun run3 = thirdPara.createRun();
	        run3.setText(docxDto.getThirdPara());
	        run3.setStrikeThrough(true);
	        run3.setColor("1607E9");
	        run3.setUnderline(UnderlinePatterns.SINGLE);
	        
	        try {
				document.write(this.out);
				//Close document
				this.out.close();
				
			} catch (IOException e) {
				logger.error(e.getMessage());
				e.printStackTrace();
			}

				
		} finally {
			try {
				document.close();
			} catch (IOException e) {
				logger.error(e.getMessage());
				e.printStackTrace();
			}
		}
		
		
		
		
		return getFileName();
	}

	@Override
	public Resource downloadFile(String fileName) {
		
		Path path=Paths.get(this.getFilePath()+"\\"+fileName);
		Resource resource=null;
		try {
			resource=new UrlResource(path.toUri());
		} catch (MalformedURLException e) {
			logger.error(e.getMessage());
			e.printStackTrace();
		}
		
		return resource;
	}

	@Override
	public boolean fileExist(String fileName) {
		boolean fileExist=true;
		
		if(this.myFile==null || !this.getFileName().equals(fileName))
			fileExist=false;
		return fileExist;
	}

}
