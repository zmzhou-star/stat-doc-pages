package com.github.zmzhoustar;

import java.io.BufferedReader;
import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.io.InputStreamReader;

import org.apache.poi.hslf.usermodel.HSLFSlideShow;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.xslf.usermodel.XMLSlideShow;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.poi.xwpf.usermodel.XWPFDocument;

import com.itextpdf.text.pdf.PdfReader;
import com.jacob.activeX.ActiveXComponent;
import com.jacob.com.Dispatch;
import com.jacob.com.Variant;

/**
 * 统计一个文件夹下所有word文档页数
 *
 * @author zmzhou
 * @version 1.0
 * date 2021/4/16 14:46
 */
public class App {

	public static void main(String[] args) throws IOException {
//		String filepath = "C:\tmp\demo"
		BufferedReader bf = new BufferedReader(new InputStreamReader(System.in));
		String path;
		System.out.println("请输入文件夹路径（ctrl+c中止，分析完成回车退出）：");
		while (!"".equals(path = bf.readLine())) {
			long total = statFileNum(path);
			System.out.println("分析完成，总页数为：" + total + "\n请输入文件夹路径（ctrl+c中止，分析完成回车退出）：");
		}
	}

	/**
	 * 统计文件夹下文档页数
	 *
	 * @param path 文件夹路径
	 * @return 页数
	 * @author zmzhou
	 * date 2021/4/16 14:42
	 */
	public static long statFileNum(String path) {
		//File类型可以是文件也可以是文件夹
		File file = new File(path);
		//将该目录下的所有文件放置在一个File类型的数组中
		File[] fileList = file.listFiles();
		long total = 0;
		assert fileList != null;
		for (File f : fileList) {
			if (f.isDirectory()) {
				// 统计子文件夹
				total += statFileNum(f.getAbsolutePath());
			} else {
				int ye = getFilePageNum(f.getAbsolutePath());
				total += ye;
				System.out.println("文件：" + f + " 页数为：" + ye);
			}
		}
		System.out.println("文件夹：" + path + " 内文档总页数为：" + total);
		return total;
	}

	/**
	 * 判断文件类型，并返回页数
	 * @param filePath 文件完整路径
	 * @return 文件页数
	 * @author zmzhou
	 * date 2021/4/21 14:16
	 */
	private static int getFilePageNum(String filePath) {
		int pageNum = 0;
		try (FileInputStream is = new FileInputStream(filePath)) {
			if (filePath.endsWith(Constants.DOC) || filePath.endsWith(Constants.DOCX)
					|| filePath.endsWith(Constants.WPS)) {
				//采用如下方法
				pageNum = getDocPageNum(filePath);
				if (filePath.endsWith(Constants.DOCX)) {
					XWPFDocument docx = new XWPFDocument(is);
					pageNum = Math.max(pageNum,
							docx.getProperties().getExtendedProperties().getUnderlyingProperties().getPages());
				}
			} else if (filePath.endsWith(Constants.PPT)) {
				HSLFSlideShow slideShow = new HSLFSlideShow(is);
				pageNum = slideShow.getSlides().size();
			} else if (filePath.endsWith(Constants.PPTX)) {
				XMLSlideShow xslideShow = new XMLSlideShow(is);
				pageNum = xslideShow.getSlides().size();
			} else if (filePath.endsWith(Constants.PDF)) {
				PdfReader reader = new PdfReader(filePath);
				pageNum = reader.getNumberOfPages();
			} else if (filePath.endsWith(Constants.XLS)) {
				HSSFWorkbook workbook = new HSSFWorkbook(is);
				int sheetNums = workbook.getNumberOfSheets();
				for (int i = 0; i < sheetNums; i++) {
					// 分页符数量
					pageNum += workbook.getSheetAt(i).getRowBreaks().length + 1;
				}
			} else if (filePath.endsWith(Constants.XLSX)) {
				XSSFWorkbook xwb = new XSSFWorkbook(is);
				int sheetNums = xwb.getNumberOfSheets();
				for (int i = 0; i < sheetNums; i++) {
					// 分页符数量
					pageNum += xwb.getSheetAt(i).getRowBreaks().length + 1;
				}
			}
		} catch (IOException e) {
			System.err.println(filePath + "，分析异常：" + e);
		}
		return pageNum;
	}


	/**
	 * 用于判断Office 2003版本之前的Word（格式为.doc）和.wps格式文档的页数
	 * @param filePath 文件完整路径
	 * @return 文件页数
	 * @author zmzhou
	 * date 2021/4/21 14:17
	 */
	private static int getDocPageNum(String filePath) {
		// 建立ActiveX部件
		ActiveXComponent wordCom = new ActiveXComponent("Word.Application");
		//word应用程序不可见
		wordCom.setProperty("Visible", false);
		// 返回wrdCom.Documents的Dispatch
		//Documents表示word的所有文档窗口（word是多文档应用程序）
		Dispatch wrdDocs = wordCom.getProperty("Documents").toDispatch();

		// 调用wrdCom.Documents.Open方法打开指定的word文档，返回wordDoc
		Dispatch wordDoc = Dispatch.call(wrdDocs, "Open", filePath, false, true, false).toDispatch();
		Dispatch selection = Dispatch.get(wordCom, "Selection").toDispatch();
		//总页数 //显示修订内容的最终状态
		int pageNum = Integer.parseInt(Dispatch.call(selection, "information", 4).toString());

		//关闭文档且不保存
		Dispatch.call(wordDoc, "Close", new Variant(false));
		//退出进程对象
		wordCom.invoke("Quit", new Variant[]{});
		return pageNum;
	}
}
