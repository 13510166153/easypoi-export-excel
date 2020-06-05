package org.example.utils;

import java.io.BufferedReader;
import java.io.StringReader;

import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.VerticalAlignment;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.ss.util.RegionUtil;
import org.apache.poi.xssf.usermodel.XSSFRichTextString;

import lombok.extern.slf4j.Slf4j;

/**
 * XLS工具类
 */
@Slf4j
public class XlsUtil {

	/**
	 * 创建字体
	 * @param name 字体名称
	 * @param size 字体尺寸
	 * @param bold 是否加粗
	 * @param workbook 工作本
	 * @return 创建的字体
	 */
	public static Font createFont(String name, short size, boolean bold, Workbook workbook) {
		Font font = workbook.createFont();
		font.setFontName(name);
		font.setFontHeightInPoints(size);
		font.setBold(bold);
		return font;
	}

	/**
	 * 创建字体
	 * @param name 字体名称
	 * @param size 字体尺寸
	 * @param bold 是否加粗
	 * @param color 字体颜色 {@link IndexedColors}
	 * @param workbook 工作本
	 * @return 创建的字体
	 */
	public static Font createFont(String name, short size, boolean bold, IndexedColors color, Workbook workbook) {
		Font font = createFont(name, size, bold, workbook);
		font.setColor(color.index);
		return font;
	}

	/**
	 * 创建带边框的单元格默认样式()
	 * @param font 字体
	 * @param size 字体尺寸
	 * @param bold 是否加粗
	 * @param color 字体颜色 {@link IndexedColors}
	 * @param workbook 工作本
	 * @return 创建的字体
	 */
	public static CellStyle createBorderCellStyle(Font font, Workbook workbook) {
		CellStyle style = workbook.createCellStyle();
		style.setAlignment(HorizontalAlignment.CENTER);// 水平
		style.setVerticalAlignment(VerticalAlignment.CENTER);// 垂直
		style.setBorderBottom(BorderStyle.THIN); // 下边框
		style.setBorderLeft(BorderStyle.THIN);// 左边框
		style.setBorderTop(BorderStyle.THIN);// 上边框
		style.setBorderRight(BorderStyle.THIN);// 右边框
		style.setWrapText(true);// 文本自动折行
		style.setFont(font);
		return style;
	}

	/**
	 * 创建单元格（XSS单元格）
	 * @param columnIndex 列索引
	 * @param style 单元格样式
	 * @param string 单元格数据(字符串)
	 * @param row 行对象
	 * @return 创建的单元格
	 */
	public static Cell createCell(int columnIndex, CellStyle style, String string, Row row) {
		Cell cell = row.createCell(columnIndex);
		cell.setCellStyle(style);
		cell.setCellValue(new XSSFRichTextString(string));
		return cell;
	}

	/**
	 * 设置合并单元格的边框
	 * @param cra 合并地址
	 * @param sheet 工作单对象
	 */
	public static void setBorder(CellRangeAddress cra, Sheet sheet) {
		RegionUtil.setBorderBottom(BorderStyle.THIN, cra, sheet); // 下边框
		RegionUtil.setBorderLeft(BorderStyle.THIN, cra, sheet); // 左边框
		RegionUtil.setBorderRight(BorderStyle.THIN, cra, sheet); // 有边框
		RegionUtil.setBorderTop(BorderStyle.THIN, cra, sheet); // 上边框
	}

	/**
	 * 计算文本行数（预估数值，不能不保证完全准确）
	 * @param text 文本字符串
	 * @param wordPerLine 每行预估字符数
	 * @return 文本行数
	 */
	public static int getCountLines(String text, int wordPerLine) {
		int count = 0;
		try (BufferedReader reader = new BufferedReader(new StringReader(text))) {
			for (String line = reader.readLine(); line != null; line = reader.readLine()) {
				int length = line.length();
				count += (length + wordPerLine - 1) / wordPerLine;
			}
		} catch (Exception e) {
			log.error("!", e);
		}
		return Math.max(count, 1);
	}
}
