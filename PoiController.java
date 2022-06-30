package com.example.demo.controller;

import java.io.File;
import java.io.IOException;

import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.web.bind.annotation.CrossOrigin;
import org.springframework.web.bind.annotation.PostMapping;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RequestParam;
import org.springframework.web.bind.annotation.RestController;
import org.springframework.web.multipart.MultipartFile;

import com.example.demo.service.PoiService;


@RestController
@RequestMapping("/")
@CrossOrigin(origins = "*",maxAge = 3600)
public class PoiController {
	
	@Autowired
	private PoiService poiservice;
	
	@RequestMapping(value = "/question")
	public String question(@RequestParam(value = "text") String keyword, @RequestParam(value = "wordpath") String wordpath) {
		String res = "结果出错!";
		try {
			res = poiservice.question(keyword, wordpath);
		} catch (IOException e) {
			e.printStackTrace();
		}
		return res;
	}
	
	@CrossOrigin
	@RequestMapping("/uploadFile")
	public String uploadFile(@RequestParam("file") MultipartFile file,@RequestParam("text") String text) {
		String result = "";
		String originalFileName = file.getOriginalFilename();
		  //logger.info("FFF"+originalFileName);
		  String fileName = "b"+System.currentTimeMillis()+"."+originalFileName.substring(originalFileName.lastIndexOf(".")+1);
		  String filePath = "D:\\www\\";
		  File fileDest = new File(filePath+fileName);
		  if(!fileDest.getParentFile().exists()) 
		   fileDest.getParentFile().mkdir();
		  try {
		   file.transferTo(fileDest);
		  }catch(Exception e) {
		   e.printStackTrace();
		  }
		  try {
			result = this.poiservice.question(text, filePath+fileName);
		} catch (IOException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
		return result;
	}
	
	@RequestMapping("/testquestion")
	public String testQuestion(){
		return "输入的问题为:";
	}
	

}
