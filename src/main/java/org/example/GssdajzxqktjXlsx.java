package org.example;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.OutputStream;
import java.util.ArrayList;
import java.util.Date;
import java.util.HashSet;
import java.util.List;
import java.util.Set;

import org.example.utils.XlsUtil;
import org.apache.commons.io.IOUtils;
import org.apache.commons.lang.time.DateFormatUtils;
import org.apache.commons.lang3.StringUtils;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;


import lombok.Data;
import lombok.experimental.Accessors;

/**
 * 各省试点案件咨询情况统计
 */
public class GssdajzxqktjXlsx {


	public static void main(String[] args) throws IOException {
		File file = new File("D:\\民行\\统计功能\\各省试点案件咨询情况统计.xlsx");
		FileOutputStream fos = new FileOutputStream(file);
		BookVo bookVo = getBookVo();
		GssdajzxqktjXlsx gssdajzxqktjXlsx=new GssdajzxqktjXlsx();
		gssdajzxqktjXlsx.render(bookVo,fos);

	}

	private static BookVo getBookVo() {
		BookVo bookVo = new BookVo();
		List<SheetVo> sheetVos = new ArrayList<>();
		for (int i = 0; i < 5; i++) {
			SheetVo sheetVo = new SheetVo();
			sheetVo.setMc("测试");
			List<TjVo> tjVos = listTjVo();
			List<XqVo> xqVos = listXqVo();
			sheetVo.setTjList(tjVos);
			sheetVo.setXqList(xqVos);
			sheetVos.add(sheetVo);
		}
		bookVo.setSheetList(sheetVos);
		return bookVo;
	}

	private static List<XqVo> listXqVo() {
		List<XqVo> xqVos = new ArrayList<>();
		for (int i = 0; i < 4; i++) {
			XqVo xqVo = new XqVo();
			xqVo.setAjmc("ajmc" + i);
			xqVo.setAjzt("ajzt" + i);
			xqVo.setCbr("cbr" + i);
			xqVo.setFbsj(new Date());
			xqVo.setZjList(listZj());
			xqVos.add(xqVo);
		}
		return xqVos;
	}

	private static List<ZjVo> listZj() {
		List<ZjVo> zjVos = new ArrayList<>();
		for (int i = 0; i < 3; i++) {
			ZjVo zjVo = new ZjVo();
			zjVo.setMc("mc" + i);
			zjVo.setZt(i);
			zjVos.add(zjVo);
		}
		return zjVos;
	}

	private static List<TjVo> listTjVo() {
		List<TjVo> tjVos = new ArrayList<>();
		for (int i = 0; i < 4; i++) {
			TjVo tjVo = new TjVo();
			tjVo.setAjzxtqqkms("案件咨询总体情况描述" + i);
			tjVo.setTjrq(new Date());
			tjVos.add(tjVo);
		}
		return tjVos;
	}

	// ==============================Fields===========================================
	private static final int CHARACTER_WIDTH = 256;

	// ==============================Methods==========================================
	/**
	 * 绘制工作表
	 * @param model 数据模型
	 * @param output 输入流
	 * @throws IOException 如果出现IO异常抛出
	 */
	public void render(BookVo model, OutputStream output) throws IOException {
		XSSFWorkbook wb = null;
		try {
			wb = new XSSFWorkbook();

			Set<String> sheelNameSet = new HashSet<>();

			for (SheetVo sheetModel : model.getSheetList()) {
				// 如果sheet名称重复，加上编号
				String mc = StringUtils.defaultIfEmpty(sheetModel.getMc(), "_");
				if (sheelNameSet.contains(mc)) {
					mc = mc + "_" + sheelNameSet.size();
				}
				sheelNameSet.add(mc);
				sheetModel.setMc(mc);

				buildSheet(sheetModel, wb);
			}

			wb.write(output);
			output.flush();
		} finally {
			IOUtils.closeQuietly(wb);
		}
	}

	private void buildSheet(SheetVo model, XSSFWorkbook wb) {

		// 获得工作表
		Sheet sheet = wb.createSheet(model.mc);

		// 设置列宽
		sheet.setColumnWidth(0, CHARACTER_WIDTH * 17);
		sheet.setColumnWidth(1, CHARACTER_WIDTH * 56);
		sheet.setColumnWidth(2, CHARACTER_WIDTH * 22);
		sheet.setColumnWidth(3, CHARACTER_WIDTH * 85);
		sheet.setColumnWidth(4, CHARACTER_WIDTH * 25);
		sheet.setColumnWidth(5, CHARACTER_WIDTH * 25);

		// 行索引
		int rowIndex = 0;

		// 合并的单元格
		List<CellRangeAddress> craList = new ArrayList<>();

		// 统计表格
		{
			Font font = XlsUtil.createFont("宋体", (short) 14, true, wb);
			CellStyle hStyle = XlsUtil.createBorderCellStyle(font, wb);

			CellStyle rStyle = XlsUtil.createBorderCellStyle(font, wb);
			rStyle.setAlignment(HorizontalAlignment.LEFT);// 水平

			// 表头
			{
				Row row = sheet.createRow(rowIndex);
				XlsUtil.createCell(0, hStyle, "统计日期", row);
				XlsUtil.createCell(1, hStyle, "案件咨询总体情况描述", row);
				sheet.addMergedRegion(new CellRangeAddress(rowIndex, rowIndex, 1, 1 + 3));
				rowIndex++;
			}

			// 主体
			for (TjVo tjVo : model.getTjList()) {
				Row row = sheet.createRow(rowIndex);
				int maxCountLines = 1;
				XlsUtil.createCell(0, hStyle, DateFormatUtils.format(tjVo.tjrq, "yyyy-MM-dd"), row);

				String ajzxtqqkms = tjVo.getAjzxtqqkms();
				XlsUtil.createCell(1, rStyle, ajzxtqqkms, row);

				// 文本可能占据的行数
				maxCountLines = XlsUtil.getCountLines(ajzxtqqkms, 70);

				// 合并单元格
				craList.add(new CellRangeAddress(rowIndex, rowIndex, 1, 4));

				row.setHeightInPoints((short) (maxCountLines * 20));
				rowIndex++;
			}

		}

		// 详情表格
		{
			CellStyle hStyle = XlsUtil.createBorderCellStyle(XlsUtil.createFont("宋体", (short) 11, false, wb), wb);
			CellStyle gStyle = XlsUtil.createBorderCellStyle(XlsUtil.createFont("宋体", (short) 11, false, IndexedColors.GREEN, wb), wb);
			CellStyle yStyle = XlsUtil.createBorderCellStyle(XlsUtil.createFont("宋体", (short) 11, false, IndexedColors.ORANGE, wb), wb);
			CellStyle rStyle = XlsUtil.createBorderCellStyle(XlsUtil.createFont("宋体", (short) 11, false, IndexedColors.RED1, wb), wb);

			// 表头
			{
				Row row = sheet.createRow(rowIndex);
				XlsUtil.createCell(0, hStyle, "序号", row);
				XlsUtil.createCell(1, hStyle, "案件名称", row);
				XlsUtil.createCell(2, hStyle, "案件状态", row);
				XlsUtil.createCell(3, hStyle, "在办专家名单（绿色表示专家已接受，黄色表示等待专家接受，红色表示专家超时未接受）", row);
				XlsUtil.createCell(4, hStyle, "案件发布时间", row);
				XlsUtil.createCell(5, hStyle, "承办人和电话号码", row);
				rowIndex++;
			}
			// 主体
			{
				int no = 1;
				for (XqVo xqVo : model.getXqList()) {

					List<ZjVo> zjList = xqVo.getZjList();
					// 行数
					int mergedCount = zjList.size();
					{
						Row row = sheet.createRow(rowIndex);
						// #0 序号
						XlsUtil.createCell(0, hStyle, String.valueOf(no++), row);
						// #1 案件名称
						XlsUtil.createCell(1, hStyle, xqVo.getAjmc(), row);

						// #2 案件状态
						XlsUtil.createCell(2, hStyle, xqVo.getAjzt(), row);

						// #3 在办专家名单（绿色表示专家已接受，黄色表示等待专家接受，红色表示专家超时未接受）
						if (mergedCount > 0) {
							ZjVo zjVo = zjList.get(0);
							int zt = zjVo.getZt();
							XlsUtil.createCell(3, //
									zt == 1 ? gStyle : // 绿
											zt == 2 ? yStyle : // 黄
													zt == 3 ? rStyle : hStyle// 红
									, zjVo.getMc(), row);
						}
						// #4 案件发布时间
						XlsUtil.createCell(4, hStyle, DateFormatUtils.format(xqVo.getFbsj(), "yyyy-MM-dd HH:mm:ss"), row);
						// #5 承办人和电话号码
						XlsUtil.createCell(5, hStyle, xqVo.getCbr(), row);

						// 合并单元格
						craList.add(new CellRangeAddress(rowIndex, rowIndex + mergedCount - 1, 0, 0));
						craList.add(new CellRangeAddress(rowIndex, rowIndex + mergedCount - 1, 1, 1));
						craList.add(new CellRangeAddress(rowIndex, rowIndex + mergedCount - 1, 2, 2));
						craList.add(new CellRangeAddress(rowIndex, rowIndex + mergedCount - 1, 4, 4));
						craList.add(new CellRangeAddress(rowIndex, rowIndex + mergedCount - 1, 5, 5));

						rowIndex++;
					}

					for (int i = 1; i < mergedCount; i++) {
						Row row = sheet.createRow(rowIndex);
						ZjVo zjVo = zjList.get(i);
						int zt = zjVo.getZt();
						XlsUtil.createCell(3, //
								zt == 1 ? gStyle : // 绿
										zt == 2 ? yStyle : // 黄
												zt == 3 ? rStyle : hStyle// 红
								, zjVo.getMc(), row);
						rowIndex++;
					}
				}
			}
		}

		for (CellRangeAddress cra : craList) {
			sheet.addMergedRegion(cra);
			XlsUtil.setBorder(cra, sheet);
		}

	}

	// ==============================InnerClass=======================================
	/**
	 * 各省试点案件咨询情况统计数据
	 */
	@Data
	@Accessors(chain = true)
	public static class BookVo {
		private List<SheetVo> sheetList = new ArrayList<>();
	}

	/** 页签 */
	@Data
	@Accessors(chain = true)
	public static class SheetVo {
		private String mc;
		private List<TjVo> tjList = new ArrayList<>();
		private List<XqVo> xqList = new ArrayList<>();
	}

	/** 统计项 */
	@Data
	@Accessors(chain = true)
	public static class TjVo {
		/** 统计日期 */
		private Date tjrq;
		/** 案件咨询总体情况描述 */
		private String ajzxtqqkms;
	}

	/** 详情项 */
	@Data
	@Accessors(chain = true)
	public static class XqVo {
		/** 案件名称 */
		private String ajmc;
		/** 案件状态 */
		private String ajzt;
		/** 在办专家名单 */
		private List<ZjVo> zjList;
		/** 发布时间 */
		private Date fbsj;
		/** 承办人和电话号码 */
		private String cbr;
	}

	/** 专家信息 */
	@Data
	@Accessors(chain = true)
	public static class ZjVo {
		/** 专家名称 */
		private String mc;
		/** 1已接受，2等待，3超时未接受 */
		private int zt;
	}
}
