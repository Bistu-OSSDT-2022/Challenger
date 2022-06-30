package com.example.demo.service;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.io.InputStream;
import java.util.ArrayList;
import java.util.Collections;
import java.util.HashMap;
import java.util.Iterator;
import java.util.List;
import java.util.Map;
import java.lang.*;

import org.apache.commons.collections4.CollectionUtils;
import org.apache.poi.POIXMLDocument;
import org.apache.poi.hwpf.extractor.WordExtractor;
import org.apache.poi.xwpf.model.XWPFHeaderFooterPolicy;
import org.apache.poi.xwpf.usermodel.IBodyElement;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFPicture;
import org.apache.poi.xwpf.usermodel.XWPFPictureData;
import org.apache.poi.xwpf.usermodel.XWPFRun;
import org.apache.poi.xwpf.usermodel.XWPFStyle;
import org.apache.poi.xwpf.usermodel.XWPFStyles;
import org.apache.poi.xwpf.usermodel.XWPFTable;
import org.apache.poi.xwpf.usermodel.XWPFTableCell;
import org.apache.poi.xwpf.usermodel.XWPFTableRow;
import org.openxmlformats.schemas.drawingml.x2006.picture.CTPicture;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTDecimalNumber;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTDocument1;
import org.springframework.stereotype.Service;


import com.example.demo.im.*;

@Service
public class PoiService {
	
	public String question(String keyword, String wordPath) throws IOException {

        String result = "";
//       return keyword;
        
        List<String> wordList1 = Segment.getWords(keyword);//get words

//       return keyword;
        Word2Vec vec = new Word2Vec();//build a model instance
        try {
            vec.loadGoogleModel("model/Google_word2vec_zhwiki2103_300d.bin");//lord dict
        } catch (IOException e) {
            e.printStackTrace();
        }
//        
//       
//        //考虑所有分词后的单词
//        List<String> wait_words = {"页眉","页脚","边距","标题","表格","段落","字数"};

        List<String> wait_words = new ArrayList<String>();
        wait_words.add("页眉");
        wait_words.add("页脚");
        wait_words.add("边距");
        wait_words.add("标题");
        wait_words.add("表格");
        wait_words.add("段落");
        wait_words.add("字数");
        wait_words.add("行数");
        wait_words.add("图片");
        wait_words.add("属性");
        wait_words.add("页数");
        
        
        int idx_max = 0;
        float max = 0;
        for(int i=0;i<wait_words.size();i++){
//            System.out.print(wait_words.get(i)+'\t');
            float tmp = vec.calMaxSimilarity(wait_words.get(i), wordList1);
            System.out.println("词最大相似度："+wait_words.get(i)+'\t'+Float.toString(tmp));
//            System.out.println(tmp);
//            ress = ress+Float.toString(max);
            if(tmp>max) {
            	idx_max = i;
            	max = tmp;
            	System.out.println("更新全局最大相似度："+wait_words.get(i)+'\t'+Float.toString(max));
            }
            //System.out.println();
        }
        System.out.println("最终全局最大相似度："+wait_words.get(idx_max)+'\t'+Float.toString(max));
        
        
        List<String> template = new ArrayList<String>();
        template.add("这篇文章的页眉是什么");//0
        template.add("这篇文章的页脚是什么");//1
        template.add("这篇文章的边距是什么");//2
        template.add("这篇文章的标题是什么");//3
        template.add("这篇文章的表格是什么");//4
        template.add("这篇文章的段落是什么");//5
        template.add("这篇文章的字数是什么");//6
        template.add("这篇文章的行数是什么");//7
        template.add("这篇文章的图片是什么");//8
        template.add("这篇文章的属性是什么");//9
        template.add("这篇文章的页数是什么");//10
        
        
        List<String> template_1 = new ArrayList<String>();
        template_1.add("页眉是什么");//0
        template_1.add("页脚是什么");//1
        template_1.add("距离是什么");//2
        template_1.add("标题是什么");//3
        template_1.add("表格是什么");//4
        template_1.add("段落是什么");//5
        template_1.add("字数是什么");//6
        template_1.add("行数是什么");//7
        template_1.add("图片是什么");//8
        template_1.add("属性是什么");//9
        template_1.add("页数是什么");//10
        
        
        float max_sen=0;
        int idx_sen = 0;
        for(int i=0;i<template.size();i++) {
        	float senten_dis=0;
        	for(int j =0;j<template.size();j++) {
        		if(senten_dis < vec.fastSentenceSimilarity(wordList1, Segment.getWords(template.get(j)))){
        			senten_dis=vec.fastSentenceSimilarity(wordList1, Segment.getWords(template.get(j)));
        		}
        	}
        	float senten_dis_1=0;
        	for(int j =0;j<template_1.size();j++) {
        		if(senten_dis_1<vec.fastSentenceSimilarity(wordList1, Segment.getWords(template_1.get(j)))){
        			senten_dis_1=vec.fastSentenceSimilarity(wordList1, Segment.getWords(template_1.get(j)));
        			//System.out.println("log");
        		}
        	}
        	System.out.println("句子1最大相似度："+template.get(i)+'\t'+Float.toString(senten_dis));
        	System.out.println("句子2最大相似度："+template_1.get(i)+'\t'+Float.toString(senten_dis_1));
        	//float senten_dis = vec.fastSentenceSimilarity(wordList1, Segment.getWords(template.get(i)));//1
        	//float senten_dis_1 = vec.fastSentenceSimilarity(wordList1, Segment.getWords(template_1.get(i)));//2
        	if(senten_dis<senten_dis_1) {
        		senten_dis=senten_dis_1;
        		System.out.println("句子2的相似度更大，相似的文本是"+template_1.get(i)+'\t'+Float.toString(senten_dis_1));
        		if(senten_dis_1>max_sen) {
            		max_sen = senten_dis_1;
            		idx_sen = i;
            		System.out.println("更新全局句子最大相似度："+template_1.get(i)+'\t'+Float.toString(max_sen));
            		//System.out.println("最终全局句子最大相似度："+template_1.get(idx_sen)+'\t'+Float.toString(max_sen));
            	}
        	}
        	else {
        		System.out.println("句子1的相似度更大，相似的文本是"+template.get(i)+'\t'+Float.toString(senten_dis));
        		if(senten_dis>max_sen) {
            		max_sen = senten_dis;
            		idx_sen = i;
            		System.out.println("更新全局句子最大相似度："+template.get(i)+'\t'+Float.toString(max_sen));
            		
            	}
        	}
        	//System.out.println("句子最大相似度："+template.get(i)+'\t'+Float.toString(senten_dis));
//        	if(senten_dis>max_sen) {
//        		max_sen = senten_dis;
//        		idx_sen = i;
//        		System.out.println("更新全局句子最大相似度："+template.get(i)+'\t'+Float.toString(max_sen));
//        	}
        }
        System.out.println("最终全局句子最大相似度："+Float.toString(max_sen));
        
        if (max<max_sen) {
        	max = max_sen;
        	System.out.println("最终取得句子");
        }
        else {
        	System.out.println("最终取得词语");
        }
        
    	if(idx_max==0 && max>0.5) {
    		result = getDefaultHeader(wordPath);
    	}
    	else if(idx_max==1 && max>0.5) {
    		result = getDefaultFooter(wordPath);
    	}
    	else if(idx_max==2 && max>0.5) {
    		result = getPgMar(wordPath);
    	}
    	else if(idx_max==3 && max>0.5) {
    		result = getParas(wordPath);
    		//获取标题
           //List<Map<String, String>> list = getParagraph(paras.get(0));
           //result="标题信息"+list;
           //System.out.println("***标题信息***\n"+list);
    	}
    	else if(idx_max==4 && max>0.5) {
    		result = getTable(wordPath);
    	}
    	else if(idx_max==5 && max>0.5) {
    		result = getPara(wordPath);	
    	}
    	else if(idx_max==6 && max>0.5) {
    		result = getWordNum(wordPath);
    	}
    	else if(idx_max==7 && max>0.5) {
    		result = getCellNum(wordPath);
    	}
    	else if(idx_max==8 && max>0.5) {
    		result = readImageInParagraph(wordPath);
    	}
    	else if(idx_max==9 && max>0.5) {
    		result = getProperty(wordPath);
    	}
    	else if(idx_max==10 && max>0.5) {
    		result = getPage(wordPath);
    	}
    	else {
    		result = "查询失败，请重新输入查询问句";
    	}
    	return result;
	}
	
	
	public static String getDefaultHeader(String wordPath) throws IOException {			//页眉
		XWPFDocument doc = new XWPFDocument(POIXMLDocument.openPackage(wordPath));
        XWPFHeaderFooterPolicy headerFooterPolicy = doc.getHeaderFooterPolicy();
        CTDocument1 ctdoc =  doc.getDocument();
      //  List<XWPFParagraph> paras = doc.getParagraphs(); //将得到包含段落列表
    	return "页眉为"+headerFooterPolicy.getDefaultHeader().getText();
    }
	
	
	public static String getDefaultFooter(String wordPath) throws IOException {			//页脚
		XWPFDocument doc = new XWPFDocument(POIXMLDocument.openPackage(wordPath));
        XWPFHeaderFooterPolicy headerFooterPolicy = doc.getHeaderFooterPolicy();
        CTDocument1 ctdoc =  doc.getDocument();
      //  List<XWPFParagraph> paras = doc.getParagraphs(); 
		return "页脚为"+headerFooterPolicy.getDefaultFooter().getText();
	}
	
	
	public static String getPgMar(String wordPath) throws IOException {				//边距
		XWPFDocument doc = new XWPFDocument(POIXMLDocument.openPackage(wordPath));
        XWPFHeaderFooterPolicy headerFooterPolicy = doc.getHeaderFooterPolicy();
        CTDocument1 ctdoc =  doc.getDocument();
        //List<XWPFParagraph> paras = doc.getParagraphs(); 
		return "上边距"+ctdoc.getBody().getSectPr().getPgMar().getTop()+"\n下边距"+ctdoc.getBody().getSectPr().getPgMar().getBottom()+"\n左边距"+ctdoc.getBody().getSectPr().getPgMar().getLeft()+"\n右边距"+ctdoc.getBody().getSectPr().getPgMar().getRight();
	}
	
	
	public static String getParas(String wordPath) throws IOException {			//标题信息
		XWPFDocument doc = new XWPFDocument(POIXMLDocument.openPackage(wordPath));
	//	List<XWPFParagraph> paragraphs = doc.getParagraphs();
        CTDocument1 ctdoc =  doc.getDocument();
        List<XWPFParagraph> paras = doc.getParagraphs(); 
		return "标题信息"+getParagraph(paras.get(0));
	}
	
	
	public static String getTable(String wordPath) throws IOException {			//表格
		XWPFDocument doc = new XWPFDocument(POIXMLDocument.openPackage(wordPath));
        XWPFHeaderFooterPolicy headerFooterPolicy = doc.getHeaderFooterPolicy();
        CTDocument1 ctdoc =  doc.getDocument();
		boolean flag=false;
		Iterator<IBodyElement> iter = doc.getBodyElementsIterator();
		String result = " ";
		while(iter.hasNext()) {
			IBodyElement element = iter.next();
			int i=getParaInformation(element);
			if(i==2) {
				flag=true;
				int row_count =0;
                XWPFTable table = (XWPFTable) element;
                List<XWPFTableRow> xwpfTableRows = table.getRows();
                row_count = xwpfTableRows.size();
                ArrayList cell_count=new ArrayList();
                int row_index = 1;
                for (XWPFTableRow xwpfTableRow : xwpfTableRows) {
                  List<XWPFTableCell> xwpfTableCells = xwpfTableRow.getTableCells();
                  cell_count.add(xwpfTableCells.size());
                  result=result+"第"+row_index+"行\n";
                  //System.out.println("第"+row_index+"行");
                  int cell_index =1;
                  for (XWPFTableCell xwpfTableCell : xwpfTableCells) {
                    //单元格是否被合并，合并了几个
                    CTDecimalNumber  cellspan = xwpfTableCell.getCTTc().getTcPr().getGridSpan();
                    boolean gridspan = cellspan != null;
                    String gridspan_num = cellspan != null?cellspan.getVal().toString():"0";
                    List<XWPFParagraph> xwpfParagraphs = xwpfTableCell.getParagraphs();
                    XWPFParagraph paragraph = xwpfParagraphs.get(0);
                    result=result+"第" +cell_index+"个单元格，合并标志："+gridspan+",合并个数:"+gridspan_num
                            +",文字："+getParagraph(paragraph);
                    //System.out.println("第" +cell_index+"个单元格，合并标志："+gridspan+",合并个数:"+gridspan_num
                    //+",文字："+getParagraph(paragraph));
                    cell_index++;
                  }
                  row_index++;
                }
                result=result+"表格为：row_cell="+row_count+"行"+Collections.max(cell_count)+"列";
                //System.out.println("表格为：row_count==="+row_count+"行"+Collections.max(cell_count)+"列");
			}
			
		}
		if(flag==false) {
			result="没有表格";
		}
		return result;
	}
	
	
	public static String getPara(String wordPath) throws IOException {		//段落
		XWPFDocument doc = new XWPFDocument(POIXMLDocument.openPackage(wordPath));
        CTDocument1 ctdoc =  doc.getDocument();
        String result=" ";
		Iterator<IBodyElement> iter = doc.getBodyElementsIterator();
		
		while(iter.hasNext()) {
			IBodyElement element = iter.next();
			int i=getParaInformation(element);
			if(i==1) {
				XWPFParagraph para = (XWPFParagraph) element;
				List<XWPFParagraph> paras = doc.getParagraphs(); 
				List<XWPFRun> runs= para.getRuns();
				XWPFParagraph para1 = null;
				String in = "";
				for (int j=0;j<paras.size(); j++) {
					
					para1 = paras.get(j);
	
					in+=getParagraph(para1).toString()+"\n";

				}
				result="段落\n"+in;
			}

		}
		
		return result;
	}
	
	
	public static String getWordNum(String wordPath) throws IOException {			//字数
		XWPFDocument doc = new XWPFDocument(POIXMLDocument.openPackage(wordPath));
        XWPFHeaderFooterPolicy headerFooterPolicy = doc.getHeaderFooterPolicy();
        CTDocument1 ctdoc =  doc.getDocument();
        List<XWPFParagraph> paras = doc.getParagraphs();
        String result = " ";
		int count = 0;
       // int i = 1;
        for (XWPFParagraph xwpfParagraph:paras) {
        	int linLength = 0;
        	String lineStr = "";
        	List<XWPFRun> xwpfRuns = xwpfParagraph.getRuns();
        	for (XWPFRun xwpfRun : xwpfRuns) {
        		linLength +=  xwpfRun.toString().trim().length();
        		lineStr += xwpfRun.toString();
				count += xwpfRun.toString().trim().length();
			}
        	//i++;
		}
        result=result+"文章总字数："+count;
        //System.out.println("文章总字数："+count);
        return result;
	}
	
	
	public static String getCellNum(String wordPath) throws IOException {			//行数
		XWPFDocument doc = new XWPFDocument(POIXMLDocument.openPackage(wordPath));
        CTDocument1 ctdoc =  doc.getDocument();
        
        String result = " ";
        
        Iterator<IBodyElement> iter = doc.getBodyElementsIterator();
		
		while(iter.hasNext()) {
			IBodyElement element = iter.next();
			int i=getParaInformation(element);
			if(i==1) {
				XWPFParagraph para = (XWPFParagraph) element;
				List<XWPFParagraph> paras = doc.getParagraphs();
				List<XWPFRun> runs= para.getRuns();
			
				result="文章总行数："+paras.size() +" 行\n";
			}
		}
        return result;
	}
	public static String readImageInParagraph(String wordPath) throws IOException {			//图片
		XWPFDocument doc = new XWPFDocument(POIXMLDocument.openPackage(wordPath));
        CTDocument1 ctdoc =  doc.getDocument();
        String result = " ";

        //List<XWPFParagraph> paras = doc.getParagraphs();

        List<XWPFParagraph> paragraphList = doc.getParagraphs();
        List<XWPFPictureData> picList = new ArrayList<>();
        boolean flag=false;
   	 	
   	 
        
        for(XWPFParagraph paragraph:paragraphList){
        	
        	List<XWPFRun> runs= paragraph.getRuns();
        	for(XWPFRun run:runs){
        		List<XWPFPicture> pictures=run.getEmbeddedPictures();
        		for(int m=0;m<pictures.size();m++){
        			if(flag=true) {
        			picList.add(pictures.get(m).getPictureData());
        			}
        			
        		}
        		result="图片信息为"+picList.toString()+"图题为"+paragraph.getParagraphText();
        	}
        		
        	}
        	

   	 	
   	 if(flag==false) {
			result="没有图片";
 	}
        return result;
        
	
	}
	 public static List<String> readImageInParagraph(XWPFParagraph paragraph) {
		 //图片索引List
		 List<String> imageBundleList = new ArrayList<String>();
		 //段落中所有XWPFRun
		 List<XWPFRun> runList = paragraph.getRuns();
		
		 
		 return imageBundleList;
		 }
		
	
	public static String getProperty(String wordPath) throws IOException {		//属性
		XWPFDocument doc = new XWPFDocument(POIXMLDocument.openPackage(wordPath));
        CTDocument1 ctdoc =  doc.getDocument();
        String result=" ";
		Iterator<IBodyElement> iter = doc.getBodyElementsIterator();
		
		while(iter.hasNext()) {
			IBodyElement element = iter.next();
			int i=getParaInformation(element);
			if(i==1) {
				XWPFParagraph para = (XWPFParagraph) element;
				List<XWPFParagraph> paras = doc.getParagraphs(); 
				List<XWPFRun> runs= para.getRuns();
			
				String in = "";
				
				in+=getParagraph1(paras.get(0)).toString();

				result=in;
			}

		}
		
		return result;
	}
	
	
	public static String getPage(String wordPath) throws IOException {		//页数
		XWPFDocument doc = new XWPFDocument(POIXMLDocument.openPackage(wordPath));
        CTDocument1 ctdoc =  doc.getDocument();
        String result=" ";
		
		int pages = doc.getProperties().getExtendedProperties().getUnderlyingProperties().getPages();//总页数
	
		result="页数="+pages;
		return result;
	}
	
	
	public static int getParaInformation(IBodyElement element) {	//判断此段落的内容是文本还是表格
    	int i=0;
    	if (element instanceof XWPFParagraph) {
            i=1;
          }
          else if (element instanceof XWPFTable) {
        	 i=2;
          }
		return i;
    	
    }
	
	
	public static List<Map<String, String>> getParagraph(XWPFParagraph para) {


        List<Map<String,String>> list = new ArrayList<>();
        Map<String,String> titile = new HashMap<>();

        titile.put("内容",para.getText());//本段全部内容


        list.add(titile);
        return list;
      }
	
	
	public static List<Map<String, String>> getParagraph1(XWPFParagraph para) {

        List<XWPFRun> runsLists = para.getRuns();
        List<Map<String,String>> list = new ArrayList<>();
        Map<String,String> titile = new HashMap<>();
        
        titile.put("对齐",para.getAlignment().toString());//对齐
        titile.put("行距",para.getSpacingBetween()+"");//行距
        titile.put("段前",para.getSpacingBeforeLines()+"");//段前
        titile.put("段后",para.getSpacingAfterLines()+"");//段后


        list.add(titile);


        for (XWPFRun run:runsLists) {

          Map<String,String> titile_map = new HashMap<>();
          //titile_map.put("content",run.getText(0));
          String Bold = Boolean.toString(run.isBold());//加粗
          titile_map.put("加粗",Bold);
          String color = run.getColor();//字体颜色
          titile_map.put("字体颜色",color);

          String FontFamily = run.getFontFamily(XWPFRun.FontCharRange.hAnsi);//字体
          titile_map.put("字体属性",FontFamily);

          String FontName = run.getFontName();//字体名
          titile_map.put("字体名",FontName);

          String FontSize = run.getFontSize()+"";//字体大小
          titile_map.put("字体大小",FontSize);

          String Underline = run.getUnderline().name();//下划线
          titile_map.put("下划线",Underline);

          String Italic =Boolean.toString(run.isItalic()) ;//字体倾斜
          titile_map.put("字体倾斜",Italic);
          list.add(titile_map);

        }
        return list;
      }
//	
//	public static int getInfor(String str) {
//		int num=0;
//		
//		char[] tochar = str.toCharArray();
//		for(int i=0;i<tochar.length-1;i++) {
//			if(tochar[i+1] == '段') {
//				num = tochar[i];
//			}
//		}
//		return num;
//	}

	
}
