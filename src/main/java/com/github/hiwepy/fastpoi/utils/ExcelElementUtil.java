/*
 * Copyright (c) 2018, hiwepy (https://github.com/hiwepy).
 *
 * Licensed under the Apache License, Version 2.0 (the "License"); you may not
 * use this file except in compliance with the License. You may obtain a copy of
 * the License at
 *
 * http://www.apache.org/licenses/LICENSE-2.0
 *
 * Unless required by applicable law or agreed to in writing, software
 * distributed under the License is distributed on an "AS IS" BASIS, WITHOUT
 * WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied. See the
 * License for the specific language governing permissions and limitations under
 * the License.
 */
package com.github.hiwepy.fastpoi.utils;

import org.apache.poi.hssf.usermodel.HSSFFooter;
import org.apache.poi.hssf.usermodel.HSSFHeader;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.Footer;
import org.apache.poi.ss.usermodel.Header;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.ss.util.CellUtil;

public class ExcelElementUtil {

	public static String getCellValue(Cell cell) {
		String value = "";
		if (cell != null) {
			
			switch (cell.getCellType()) {
			case Cell.CELL_TYPE_BOOLEAN:
				value = String.valueOf(cell.getBooleanCellValue());
				break;
			case Cell.CELL_TYPE_NUMERIC:
				value = String.valueOf(cell.getNumericCellValue());
				break;
			case Cell.CELL_TYPE_STRING:
				value = cell.getRichStringCellValue().getString();
				break;
			case Cell.CELL_TYPE_ERROR:
				value = String.valueOf(cell.getErrorCellValue());
			default:
				break;
			}
		}
		return value;
	}
	
	
	/**
	 * 
	 * @description:创建Sheet工作簿
	 * @author <a href="mailto:564039439@qq.com">wandalong</a>
	 * @date 2012-1-31
	 * @param wb
	 * @param sheetName
	 * @return
	 */
	public static Sheet createSheet(Workbook wb,String sheetName) {
		Sheet sheet = wb.createSheet(sheetName);
		sheet.setDefaultColumnWidth(12);
		sheet.setDisplayGridlines(false);
		return sheet;
	}
	
	/**
	 * 
	 * @description:创建Row
	 * @author <a href="mailto:564039439@qq.com">wandalong</a>
	 * @date 2012-1-31
	 * @param sheet Sheet
	 * @param rowNum int
	 * @param height int 行高
	 * @return HSSFRow
	 */
	public static Row createRow(Sheet sheet, int rowNum, Integer height) {
		Row row = sheet.createRow(rowNum);
		if (height != null) {
			row.setHeight(Short.parseShort(height.toString()));
		}
		return row;
	}

	/**
	 * 
	 * @description:创建CellStyle样式
	 * @author <a href="mailto:564039439@qq.com">wandalong</a>
	 * @date 2012-1-31
	 * @param wb Workbook
	 * @param backgroundColor 背景色
	 * @param foregroundColor 前置色
	 * @param halign
	 * @param font 字体
	 * @return
	 */
	public static CellStyle createCellStyle(Workbook wb,short backgroundColor, short foregroundColor, 
			short halign,Font font) {
		CellStyle cs = wb.createCellStyle();
		cs.setAlignment(halign);
		cs.setVerticalAlignment(CellStyle.VERTICAL_CENTER);
		cs.setFillBackgroundColor(backgroundColor);
		cs.setFillForegroundColor(foregroundColor);
		cs.setFillPattern(CellStyle.SOLID_FOREGROUND);
		cs.setFont(font);
		return cs;
	}

	/**
	 * 
	 * @description:创建带边框的CellStyle样式
	 * @author <a href="mailto:564039439@qq.com">wandalong</a>
	 * @date 2012-1-31
	 * @param wb Workbook
	 * @param backgroundColor 背景色
	 * @param foregroundColor 前置色
	 * @param halign
	 * @param font 字体
	 * @return
	 */
	public static CellStyle createBorderCellStyle(Workbook wb,short backgroundColor, short foregroundColor,
			short halign,Font font) {
		CellStyle cs = wb.createCellStyle();
		cs.setAlignment(halign);
		cs.setVerticalAlignment(CellStyle.VERTICAL_CENTER);
		cs.setFillBackgroundColor(backgroundColor);
		cs.setFillForegroundColor(foregroundColor);
		cs.setFillPattern(CellStyle.SOLID_FOREGROUND);
		cs.setFont(font);
		cs.setBorderLeft(CellStyle.BORDER_DASHED);
		cs.setBorderRight(CellStyle.BORDER_DASHED);
		cs.setBorderTop(CellStyle.BORDER_DASHED);
		cs.setBorderBottom(CellStyle.BORDER_DASHED);
		return cs;
	}
	
	/**
	 * 
	 * @description:创建单元格边框为四周环绕的CellStyle样式
	 * @author <a href="mailto:564039439@qq.com">wandalong</a>
	 * @date 2012-1-31
	 * @param wb Workbook
	 * @param backgroundColor 背景色
	 * @param foregroundColor 前置色
	 * @param halign
	 * @param font 字体
	 * @return
	 */
	public static CellStyle createBorderCellStyle2(Workbook wb,short backgroundColor, short foregroundColor,
			short halign,Font font) {
		CellStyle style = wb.createCellStyle(); 
	    style.setBorderBottom(CellStyle.BORDER_THIN); 
	    style.setBottomBorderColor(IndexedColors.BLACK.getIndex()); 
	    style.setBorderLeft(CellStyle.BORDER_THIN); 
	    style.setLeftBorderColor(IndexedColors.GREEN.getIndex()); 
	    style.setBorderRight(CellStyle.BORDER_THIN); 
	    style.setRightBorderColor(IndexedColors.BLUE.getIndex()); 
	    style.setBorderTop(CellStyle.BORDER_MEDIUM_DASHED); 
	    style.setTopBorderColor(IndexedColors.BLACK.getIndex()); 
		return style;
	}

	public static void setBorder(Workbook book, Integer top,Integer bottom, Integer left, Integer right) {
		CellStyle style = book.createCellStyle();
		if (top != null) {
			style.setBorderTop(Short.parseShort(top.toString()));
		}
		if (bottom != null) {
			style.setBorderBottom(Short.parseShort(bottom.toString()));
		}

		if (left != null) {
			style.setBorderLeft(Short.parseShort(left.toString()));
		}
		if (right != null) {
			style.setBorderRight(Short.parseShort(right.toString()));
		}
	}
	
	/**
	 * 
	 * @description:创建Cell
	 * @author <a href="mailto:564039439@qq.com">wandalong</a>
	 * @date 2012-1-31
	 * @param row Row
	 * @param cellNum int 列索引
	 * @param style HSSFStyle样式
	 * @return
	 */
	public static Cell createCell(Row row, int cellNum, CellStyle style) {
		Cell cell = row.createCell(cellNum);
		cell.setCellStyle(style);
		return cell;
	}

	/**
	 * 
	 * @description:合并单元格
	 * @author <a href="mailto:564039439@qq.com">wandalong</a>
	 * @date 2012-1-31
	 * @param sheet
	 * @param firstRow
	 * @param lastRow
	 * @param firstColumn
	 * @param lastColumn
	 * @return int 合并区域号码
	 */
	public static int mergeCell(Sheet sheet, int firstRow, int lastRow,int firstColumn, int lastColumn) {
		return sheet.addMergedRegion(new CellRangeAddress(firstRow, lastRow,firstColumn, lastColumn));
	}

	/**
	 * 
	 * @description:创建字体
	 * @author <a href="mailto:564039439@qq.com">wandalong</a>
	 * @date 2012-1-31
	 * @param wb Workbook
	 * @param boldweight short
	 * @param color short
	 * @param size short
	 * @return
	 */
	public static Font createFont(Workbook wb, short boldweight,short color, short size) {
		Font font = wb.createFont();
		font.setBoldweight(boldweight);
		font.setColor(color);
		font.setFontHeightInPoints(size);
		return font;
	}

	/**
	 * 
	 * @description:设置合并单元格的边框样式
	 * @author <a href="mailto:564039439@qq.com">wandalong</a>
	 * @date 2012-1-31
	 * @param sheet HSSFSheet
	 * @param ca CellRangAddress
	 * @param style CellStyle
	 */
	public static void setRegionStyle(Sheet sheet, CellRangeAddress ca,CellStyle style) {
		for (int i = ca.getFirstRow(); i <= ca.getLastRow(); i++) {
			Row row = CellUtil.getRow(i, sheet);
			for (int j = ca.getFirstColumn(); j <= ca.getLastColumn(); j++) {
				Cell cell = CellUtil.getCell(row, j);
				cell.setCellStyle(style);
			}
		}
	}

	public static Font setFont(Workbook workbook, int fontHeight,short boldWeight) {
		// 字体
		Font font = workbook.createFont();
		font.setFontName("宋体");
		font.setFontHeightInPoints((short) fontHeight);
		font.setBoldweight(boldWeight);
		return font;
	}

	public static void setSheetHeadAndFoot(Sheet sheet) {
		Header header = sheet.getHeader();
		header.setCenter("Center Header");
		header.setLeft("Left Header");
		header.setRight(HSSFHeader.font("Stencil-Normal", "Italic")+ HSSFHeader.fontSize((short) 16)+ "Right w/ Stencil-Normal Italic font and size 16");
		Footer footer = sheet.getFooter();
		footer.setRight("Page " + HSSFFooter.page() + " of "+ HSSFFooter.numPages());
	}
	
}