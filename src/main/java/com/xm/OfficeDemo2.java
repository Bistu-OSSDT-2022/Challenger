package com.xm;

import cn.hutool.core.codec.Base64;
import lombok.EqualsAndHashCode;
import org.apache.poi.xwpf.model.XWPFHeaderFooterPolicy;
import org.apache.poi.xwpf.usermodel.*;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTDecimalNumber;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTDocument1;

import java.io.FileInputStream;
import java.io.IOException;
import java.io.InputStream;
import java.util.*;

@EqualsAndHashCode
public class OfficeDemo2 {
    public static void main(String[] args) throws IOException {
            InputStream is = new FileInputStream("C:\\Users\\26428\\Desktop\\Spring 框架的复习与学习.docx");
            XWPFDocument doc = new XWPFDocument(is);
            XWPFHeaderFooterPolicy headerFooterPolicy = doc.getHeaderFooterPolicy();
            //获取页眉
            String header = headerFooterPolicy.getDefaultHeader().getText();
            System.out.println("***页眉 ***"+header);
//            获取页脚
            String footer = headerFooterPolicy.getDefaultFooter().getText();
            System.out.println("***页脚 ***"+footer);

            CTDocument1 ctdoc =  doc.getDocument();

            //System.out.println(ctdoc.getBody().getSectPr().getPgNumType().getStart());
            //上下左右边距
            System.out.println("***上边距***"+ctdoc.getBody().getSectPr().getPgMar().getTop());
            System.out.println("***下边距***"+ctdoc.getBody().getSectPr().getPgMar().getBottom());
            System.out.println("***左边距***"+ctdoc.getBody().getSectPr().getPgMar().getLeft());
            System.out.println("***右边距***"+ctdoc.getBody().getSectPr().getPgMar().getRight());

            System.out.println("*******************");

            List<XWPFParagraph> paras = doc.getParagraphs(); //将得到包含段落列表

            //获取标题
            List<Map<String, String>> list = getParagraph(paras.get(0));
            System.out.println("***标题信息**"+list);

            Iterator<IBodyElement> iter = doc.getBodyElementsIterator();
            while (iter.hasNext()) {
                // iter.next();

                IBodyElement element = iter.next();

                if (element instanceof XWPFParagraph) {
                    XWPFParagraph para = (XWPFParagraph) element;
                    System.out.println("para==="+getParagraph(para));
                }else if (element instanceof XWPFTable) {
                    int row_count =0;
                    XWPFTable table = (XWPFTable) element;
                    List<XWPFTableRow> xwpfTableRows = table.getRows();
                    row_count = xwpfTableRows.size();
                    ArrayList cell_count=new ArrayList();
                    int row_index = 1;
                    for (XWPFTableRow xwpfTableRow : xwpfTableRows) {
                        List<XWPFTableCell> xwpfTableCells = xwpfTableRow.getTableCells();
                        cell_count.add(xwpfTableCells.size());
                        System.out.println("第"+row_index+"行");
                        int cell_index =1;
                        for (XWPFTableCell xwpfTableCell : xwpfTableCells) {
                            //单元格是否被合并，合并了几个
                            CTDecimalNumber cellspan = xwpfTableCell.getCTTc().getTcPr().getGridSpan();
                            boolean gridspan = cellspan != null;
                            String gridspan_num = cellspan != null?cellspan.getVal().toString():"0";
                            List<XWPFParagraph> xwpfParagraphs = xwpfTableCell.getParagraphs();
                            XWPFParagraph paragraph = xwpfParagraphs.get(0);
                            System.out.println("第" +cell_index+"个单元格，合并标志："+gridspan+",合并个数:"+gridspan_num
                                    +"文字："+getParagraph(paragraph));
                            cell_index++;
                        }
                        row_index++;
                    }
                    System.out.println("表格为：row_count==="+row_count+"行"+ Collections.max(cell_count)+"列");
                }
            }

        }
        private static List<Map<String, String>> getParagraph(XWPFParagraph para) {
            //第一段即首行 为标题
            // XWPFParagraph para = paras.get(0);
            //标题内容
            List<XWPFRun> runsLists = para.getRuns();
            List<Map<String,String>> list = new ArrayList<Map<String, String>>();
            Map<String,String> titile = new HashMap<String, String>();
            titile.put("Text",para.getText());//本段全部内容
            titile.put("Alignment",para.getAlignment().toString());
            titile.put("SpacingBetween",para.getSpacingBetween()+"");//行距
            titile.put("SpacingBeforeLines",para.getSpacingBeforeLines()+"");//段前
            titile.put("SpacingAfterLines",para.getSpacingAfterLines()+"");//段后
            titile.put("NumLevelText",para.getNumLevelText()+"");//自动编号格式
            list.add(titile);
            //先判断缩进方式再进行数值计算
            double ind=-1,ind_left=-1,ind_right=-1,ind_hang=-1;
            String  ind_type="";
            if(para.getIndentationHanging()!=-1){//悬挂缩进
                ind_type = "hang";
                if (para.getIndentationHanging()%567 ==0 ){//悬挂单位为厘米
                    ind  = para.getIndentationHanging()/567.0;
                    ind_left = (para.getIndentationLeft()-567.0*ind)/210;
                }else{//悬挂单位为字符
                    ind  = para.getIndentationHanging()/240;
                    ind_left =  (para.getIndentationLeft()-para.getIndentationHanging())/210;
                }
                ind_right = para.getIndentationRight()/210.0;
            }else{//首行缩进或者无
                ind_type = "first";
                if(para.getFirstLineIndent() == -1){
                    ind_type = "none";
                    ind = 0;
                }else{
                    ind  = para.getFirstLineIndent()%567.0==0?para.getFirstLineIndent()/567.0:para.getFirstLineIndent()/240.0;
                }
                ind_left = para.getIndentationLeft()/210;
                ind_right = para.getIndentationRight()/210.0;
            }
            //System.out.println(ind_type+","+ind+","+ind_left+","+ind_right);
            for (XWPFRun run:runsLists
            ) {
                List<XWPFPicture> pictures = run.getEmbeddedPictures();
                if(pictures.size()>0){
                    XWPFPicture picture = pictures.get(0);
                    XWPFPictureData pictureData = picture.getPictureData();
                    //System.out.println(pictureData.getPictureType());
                    // System.out.println(picture);
                    //实现不了查询图片环绕方式
                    System.out.println(Base64.encode(pictureData.getData()));
                }
                Map<String,String> titile_map = new HashMap<String, String>();
                titile_map.put("content",run.getText(0));
                String Bold = Boolean.toString(run.isBold());//加粗
                titile_map.put("Bold",Bold);
                String color = run.getColor();//字体颜色
                titile_map.put("Color",color);
                String FontFamily = run.getFontFamily(XWPFRun.FontCharRange.hAnsi);//字体
                titile_map.put("FontFamily",FontFamily);
                String FontName = run.getFontName();//字体
                titile_map.put("FontName",FontName);
                String FontSize = run.getFontSize()+"";//字体大小
                titile_map.put("FontSize",FontSize);
                String Underline = run.getUnderline().name();//字下加线
                titile_map.put("Underline",Underline);
                String UnderlineColor = run.getUnderlineColor();//字下加线颜色
                titile_map.put("UnderlineColor",UnderlineColor);
                String Italic =Boolean.toString(run.isItalic()) ;//字体倾斜
                titile_map.put("Italic",Italic);
                list.add(titile_map);

            }
            return list;
        }
}
