package com.modeltodocx;


import java.io.File;
import java.io.FileInputStream; 
import java.io.FileNotFoundException;
import java.io.FileOutputStream; 
import java.io.IOException; 
import java.io.InputStream; 
import java.io.OutputStream; 
import java.util.ArrayList; 
import java.util.Date;
import java.util.HashMap; 
import java.util.Iterator; 
import java.util.List; 
import java.util.Map; 
import java.util.regex.Matcher; 
import java.util.regex.Pattern;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.xwpf.usermodel.Document;
import org.apache.poi.xwpf.usermodel.XWPFDocument; 
import org.apache.poi.xwpf.usermodel.XWPFParagraph; 
import org.apache.poi.xwpf.usermodel.XWPFRun; 
import org.apache.poi.xwpf.usermodel.XWPFTable; 
import org.apache.poi.xwpf.usermodel.XWPFTableCell; 
import org.apache.poi.xwpf.usermodel.XWPFTableRow; 
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTFonts; 
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTInd; 
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTP; 
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTPPr; 
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTRPr; 
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTSpacing; 
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTTc; 
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTTcPr;

public class GeneralTemplateDOCX {

	public static void main(String[] args) {
	String filePath = "D:/Work/javaweb/template/模板.docx";//模板  无表格
	String res = String.valueOf(new Date().getTime());
	String outFile = "D:/DOC/doc/模板生成文件" + res + ".docx";
    try {
        GeneralTemplateDOCX gt = new GeneralTemplateDOCX();
        Map<String, Object> params = gt.getParams();        		
        gt.templateWrite(filePath, outFile, params, gt.generateTestData(5),gt.generateTestData2(4));
        System.out.println("生成模板成功");
        System.out.println(outFile);
    } catch (Exception e) {
        System.out.println("生成模板失败");
        e.printStackTrace();
    }
}
	private Map<String, Object> getParams() {
		Map<String, Object> params = new HashMap<String, Object>();
		params.put("title","  标题文字" );
		
		params.put("myTable1","肝检查" );
		params.put("myTable2","肺检查" );     
		params.put("name", "小宝");
		params.put("age", "18");
		params.put("sex", "男");
		params.put("job", "医师");
		params.put("hobby", "电商");
		params.put("phone", "1717"); 
		params.put("name2", "2小宝");
		params.put("age2", "218");
		params.put("sex2", "2男");
		params.put("job2", "2医师");	
		return params;
	}

	// 生成tab1测试数据
	public List<List<String>> generateTestData(int num) {
	    List<List<String>> resultList = new ArrayList<List<String>>();
	    for (int i = 1; i <= num; i++) {
	        List<String> list = new ArrayList<String>();
	        list.add("" + i);
	        list.add("测试_" + i);
	        list.add("测试2_" + i);
	        list.add("测试3_" + i);
	        list.add("测试4_" + i);
	        list.add("测试5_" + i);
	        resultList.add(list);
	    }
	    return resultList;
	}
	// 生成tab2测试数据
	public List<List<String>> generateTestData2(int num) {
		List<List<String>> resultList = new ArrayList<List<String>>();
		for (int i = 1; i <= num; i++) {
			List<String> list = new ArrayList<String>();
			list.add("2_  " + i);
			list.add("2测试1_" + i);
			list.add("2测试2_" + i);
			list.add("2测试3_" + i);
//			list.add("2测试4_" + i);
//			list.add("2测试5_" + i);
			resultList.add(list);
		}
		return resultList;
	}
	
/**
 * 用一个docx文档作为模板，然后替换其中的内容，再写入目标文档中。
 * @param list 
 * 
 * @throws Exception
 */
public void templateWrite(String filePath, String outFile,
        Map<String, Object> params,List<List<String>> list1, List<List<String>> list2) throws Exception {
    InputStream is = new FileInputStream(filePath);
    System.out.println(filePath);
    XWPFDocument doc = new XWPFDocument(is); 
    // 替换段落里面的变量
    this.replaceInPara(doc, params);
    // 替换多个表格里面的变量并插入对应数据  Flag1 插入resultList Flag2 插入list 数据
    this.insertValueToTables(doc, params,list1,list2);
    OutputStream os = new FileOutputStream(outFile);
    doc.write(os);
    this.close(os);
    this.close(is);
    String imageFile ="D:/Work/条形码.jpg";
    this.insertimageToDoc(outFile,imageFile);
}

private void insertimageToDoc(String outFile, String imageFile)
		throws FileNotFoundException, IOException, InvalidFormatException {
	CustomXWPFDocument document = new CustomXWPFDocument(
			new FileInputStream(new File(outFile)));
	FileOutputStream fos = new FileOutputStream(new File(outFile));
	String blipId = document.addPictureData(new FileInputStream(new File(
			imageFile)), Document.PICTURE_TYPE_JPEG);
	document.createPicture(blipId,
			document.getNextPicNameNumber(Document.PICTURE_TYPE_JPEG), 250,
			50);
	document.write(fos);
	fos.flush();
	fos.close();
}

/**
 * 替换段落里面的变量
 * 
 * @param doc
 *            要替换的文档
 * @param params
 *            参数
 */
private void replaceInPara(XWPFDocument doc, Map<String, Object> params) {
    Iterator<XWPFParagraph> iterator = doc.getParagraphsIterator();
    XWPFParagraph para;
    while (iterator.hasNext()) {
        para = iterator.next();
        this.replaceInPara(para, params);
    }
}

/**
 * 替换段落里面的变量
 * 
 * @param para
 *            要替换的段落
 * @param params
 *            参数
 */
private boolean replaceInPara(XWPFParagraph para, Map<String, Object> params) {
    boolean data = false;
    List<XWPFRun> runs;
    Matcher matcher;
    if (this.matcher(para.getParagraphText()).find()) {
        runs = para.getRuns();
        for (int i = 0; i < runs.size(); i++) {
            XWPFRun run = runs.get(i);
            String runText = run.toString();
            System.out.println("将被替代的列值："+runText);
            matcher = this.matcher(runText);
            if (matcher.find()) {
                while ((matcher = this.matcher(runText)).find()) {
                    runText = matcher.replaceFirst(String.valueOf(params
                            .get(matcher.group(1))));
                }
                // 直接调用XWPFRun的setText()方法设置文本时，在底层会重新创建一个XWPFRun，把文本附加在当前文本后面，
                // 所以我们不能直接设值，需要先删除当前run,然后再自己手动插入一个新的run。
                para.removeRun(i);
                para.insertNewRun(i).setText(runText);
            }
        }
        data = true;
    } else if (this.matcherRow(para.getParagraphText())) {
        runs = para.getRuns();

        System.out.println("run  " + runs);

        data = true;
    }
    return data;
}

/**
 * 按模版行样式填充数据,暂未实现特殊样式填充(如列合并)，只能用于普通样式(如段落间距 缩进 字体 对齐)
 *  Flag1 插入resultList Flag2 插入list 数据
 * @param doc
 *            要替换的文档
 * @param params
 *            参数
 * @param resultList
 *            需要遍历的数据
 * @param list2 
 * @throws Exception
 */
private void insertValueToTables(XWPFDocument doc, Map<String, Object> params,List<List<String>> list1, List<List<String>> list2)
        throws Exception {
    Iterator<XWPFTable> iterator = doc.getTablesIterator();
    XWPFTable table = null;
    String tflag ="";//表标识
//    int z =1;
    //找到表 并获取第二行第一列标识名称 Flag1和Flag2
    while (iterator.hasNext()) {
//    	System.out.println("-------解析出第 "+z+" 个表------开始处理");
    	List<XWPFTableRow> rows = null;
    	List<XWPFTableCell> cells = null;
    	List<XWPFParagraph> paras;
    	XWPFTableRow tmpRow = null;// 匹配用
    	int thisRow=2;
        table = iterator.next();
		rows = table.getRows();
		for (int i = 1; i <= rows.size(); i++) {
			//获取当前行所有列
			cells = rows.get(i - 1).getTableCells();
			int intcell = 1;
			//遍历列
			for (XWPFTableCell cell : cells) {
//				System.out.println(intcell + "列：" + cell.getText()+ "  row=" + i);
				if (i == 1) {
					paras = cell.getParagraphs();
					for (XWPFParagraph para : paras) {
						// 判断是否含有${}值的列并对应替换掉
						if (this.replaceInPara(para, params)) {
							//thisRow = i;// 找到模板行定死第二行为模板行
							tmpRow = rows.get(1);
						}
					}
				} else if(i == 2){
					//取第二行第一列
					if(intcell==1) {
						tflag = cell.getText();
						break;
					}
				}
				intcell++;
			}
		}
        if("Flag1".equals(tflag)){
            this.insertValueToTable(tmpRow, table,list1,thisRow);
        }else if("Flag2".equals(tflag)){
       	    this.insertValueToTable(tmpRow, table,list2,thisRow);
        }
//        else{
//        	System.out.println("没找到table！");
//        }
//        z++;
    }
}

private void insertValueToTable(XWPFTableRow tmpRow, XWPFTable table,
		List<List<String>> list1, int thisRow) throws Exception {
	List<XWPFTableCell> tmpCells = null;// 模版列
	XWPFTableCell tmpCell = null;// 匹配用
	List<XWPFTableCell> cells = null;
	tmpCells = tmpRow.getTableCells();
	for (int i = 0, len = list1.size(); i < len; i++) {
		System.out.println("开始写第" + i + "行");
		XWPFTableRow row = table.createRow();
		row.setHeight(tmpRow.getHeight());
		List<String> list = list1.get(i);
		cells = row.getTableCells();
		// 插入的行会填充与表格第一行相同的列数
		for (int k = 0, klen = cells.size(); k < klen; k++) {
			tmpCell = tmpCells.get(k);
			XWPFTableCell cell = cells.get(k);
			setCellText(tmpCell, cell, list.get(k));
		}
		// 继续写剩余的列
		for (int j = cells.size(), jlen = list.size(); j < jlen; j++) {
			tmpCell = tmpCells.get(j);
			XWPFTableCell cell = row.addNewTableCell();
			setCellText(tmpCell, cell, list.get(j));
			System.out.println("内容" + list.get(j));
		}
	}
	// 删除模版行
	table.removeRow(thisRow - 1);
	
}

public void setCellText(XWPFTableCell tmpCell, XWPFTableCell cell,
        String text) throws Exception {
    CTTc cttc2 = tmpCell.getCTTc();
    CTTcPr ctPr2 = cttc2.getTcPr();

    CTTc cttc = cell.getCTTc();
    CTTcPr ctPr = cttc.addNewTcPr();
    cell.setColor(tmpCell.getColor());
    // cell.setVerticalAlignment(tmpCell.getVerticalAlignment());
    if (ctPr2.getTcW() != null) {
        ctPr.addNewTcW().setW(ctPr2.getTcW().getW());
    }
    if (ctPr2.getVAlign() != null) {
        ctPr.addNewVAlign().setVal(ctPr2.getVAlign().getVal());
    }
    if (cttc2.getPList().size() > 0) {
        CTP ctp = cttc2.getPList().get(0);
        if (ctp.getPPr() != null) {
            if (ctp.getPPr().getJc() != null) {
                cttc.getPList().get(0).addNewPPr().addNewJc()
                        .setVal(ctp.getPPr().getJc().getVal());
            }
        }
    }
    if (ctPr2.getTcBorders() != null) {
        ctPr.setTcBorders(ctPr2.getTcBorders());
    }
    XWPFParagraph tmpP = tmpCell.getParagraphs().get(0);
    XWPFParagraph cellP = cell.getParagraphs().get(0);
    XWPFRun tmpR = null;
    if (tmpP.getRuns() != null && tmpP.getRuns().size() > 0) {
        tmpR = tmpP.getRuns().get(0);
    }
    XWPFRun cellR = cellP.createRun();
    cellR.setText(text);
    // 复制字体信息
    if (tmpR != null) {
        cellR.setBold(tmpR.isBold());
        cellR.setItalic(tmpR.isItalic());
        cellR.setStrike(tmpR.isStrike());
        cellR.setUnderline(tmpR.getUnderline());
        cellR.setColor(tmpR.getColor());
        cellR.setTextPosition(tmpR.getTextPosition());
        if (tmpR.getFontSize() != -1) {
            cellR.setFontSize(tmpR.getFontSize());
        }
        if (tmpR.getFontFamily() != null) {
            cellR.setFontFamily(tmpR.getFontFamily());
        }
        if (tmpR.getCTR() != null) {
            if (tmpR.getCTR().isSetRPr()) {
                CTRPr tmpRPr = tmpR.getCTR().getRPr();
                if (tmpRPr.isSetRFonts()) {
                    CTFonts tmpFonts = tmpRPr.getRFonts();
                    CTRPr cellRPr = cellR.getCTR().isSetRPr() ? cellR
                            .getCTR().getRPr() : cellR.getCTR().addNewRPr();
                    CTFonts cellFonts = cellRPr.isSetRFonts() ? cellRPr
                            .getRFonts() : cellRPr.addNewRFonts();
                    cellFonts.setAscii(tmpFonts.getAscii());
                    cellFonts.setAsciiTheme(tmpFonts.getAsciiTheme());
                    cellFonts.setCs(tmpFonts.getCs());
                    cellFonts.setCstheme(tmpFonts.getCstheme());
                    cellFonts.setEastAsia(tmpFonts.getEastAsia());
                    cellFonts.setEastAsiaTheme(tmpFonts.getEastAsiaTheme());
                    cellFonts.setHAnsi(tmpFonts.getHAnsi());
                    cellFonts.setHAnsiTheme(tmpFonts.getHAnsiTheme());
                }
            }
        }
    }
    // 复制段落信息
    cellP.setAlignment(tmpP.getAlignment());
    cellP.setVerticalAlignment(tmpP.getVerticalAlignment());
    cellP.setBorderBetween(tmpP.getBorderBetween());
    cellP.setBorderBottom(tmpP.getBorderBottom());
    cellP.setBorderLeft(tmpP.getBorderLeft());
    cellP.setBorderRight(tmpP.getBorderRight());
    cellP.setBorderTop(tmpP.getBorderTop());
    cellP.setPageBreak(tmpP.isPageBreak());
    if (tmpP.getCTP() != null) {
        if (tmpP.getCTP().getPPr() != null) {
            CTPPr tmpPPr = tmpP.getCTP().getPPr();
            CTPPr cellPPr = cellP.getCTP().getPPr() != null ? cellP
                    .getCTP().getPPr() : cellP.getCTP().addNewPPr();
            // 复制段落间距信息
            CTSpacing tmpSpacing = tmpPPr.getSpacing();
            if (tmpSpacing != null) {
                CTSpacing cellSpacing = cellPPr.getSpacing() != null ? cellPPr
                        .getSpacing() : cellPPr.addNewSpacing();
                if (tmpSpacing.getAfter() != null) {
                    cellSpacing.setAfter(tmpSpacing.getAfter());
                }
                if (tmpSpacing.getAfterAutospacing() != null) {
                    cellSpacing.setAfterAutospacing(tmpSpacing
                            .getAfterAutospacing());
                }
                if (tmpSpacing.getAfterLines() != null) {
                    cellSpacing.setAfterLines(tmpSpacing.getAfterLines());
                }
                if (tmpSpacing.getBefore() != null) {
                    cellSpacing.setBefore(tmpSpacing.getBefore());
                }
                if (tmpSpacing.getBeforeAutospacing() != null) {
                    cellSpacing.setBeforeAutospacing(tmpSpacing
                            .getBeforeAutospacing());
                }
                if (tmpSpacing.getBeforeLines() != null) {
                    cellSpacing.setBeforeLines(tmpSpacing.getBeforeLines());
                }
                if (tmpSpacing.getLine() != null) {
                    cellSpacing.setLine(tmpSpacing.getLine());
                }
                if (tmpSpacing.getLineRule() != null) {
                    cellSpacing.setLineRule(tmpSpacing.getLineRule());
                }
            }
            // 复制段落缩进信息
            CTInd tmpInd = tmpPPr.getInd();
            if (tmpInd != null) {
                CTInd cellInd = cellPPr.getInd() != null ? cellPPr.getInd()
                        : cellPPr.addNewInd();
                if (tmpInd.getFirstLine() != null) {
                    cellInd.setFirstLine(tmpInd.getFirstLine());
                }
                if (tmpInd.getFirstLineChars() != null) {
                    cellInd.setFirstLineChars(tmpInd.getFirstLineChars());
                }
                if (tmpInd.getHanging() != null) {
                    cellInd.setHanging(tmpInd.getHanging());
                }
                if (tmpInd.getHangingChars() != null) {
                    cellInd.setHangingChars(tmpInd.getHangingChars());
                }
                if (tmpInd.getLeft() != null) {
                    cellInd.setLeft(tmpInd.getLeft());
                }
                if (tmpInd.getLeftChars() != null) {
                    cellInd.setLeftChars(tmpInd.getLeftChars());
                }
                if (tmpInd.getRight() != null) {
                    cellInd.setRight(tmpInd.getRight());
                }
                if (tmpInd.getRightChars() != null) {
                    cellInd.setRightChars(tmpInd.getRightChars());
                }
            }
        }
    }
}

/**
 * 正则匹配字符串
 * 
 * @param str
 * @return
 */
private Matcher matcher(String str) {
    Pattern pattern = Pattern.compile("\\$\\{(.+?)\\}",
            Pattern.CASE_INSENSITIVE);
    Matcher matcher = pattern.matcher(str);
    return matcher;
}

/**
 * 正则匹配字符串
 * 
 * @param str
 * @return
 */
private boolean matcherRow(String str) {
    Pattern pattern = Pattern.compile("\\$\\[(.+?)\\]",
            Pattern.CASE_INSENSITIVE);
    Matcher matcher = pattern.matcher(str);
    return matcher.find();
}

/**
 * 关闭输入流
 * 
 * @param is
 */
private void close(InputStream is) {
    if (is != null) {
        try {
            is.close();
        } catch (IOException e) {
            e.printStackTrace();
        }
    }
}

	/**
	 * 关闭输出流
	 * 
	 * @param os
	 */
	private void close(OutputStream os) {
	    if (os != null) {
	        try {
	            os.close();
	        } catch (IOException e) {
	            e.printStackTrace();
	        }
	    }
	}

}