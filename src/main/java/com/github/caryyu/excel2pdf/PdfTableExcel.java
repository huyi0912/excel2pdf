package com.github.caryyu.excel2pdf;

import com.itextpdf.text.*;
import com.itextpdf.text.Font;
import com.itextpdf.text.pdf.PdfPCell;
import com.itextpdf.text.pdf.PdfPTable;

import org.apache.poi.hssf.usermodel.*;
import org.apache.poi.hssf.util.HSSFColor;
import org.apache.poi.ss.format.*;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;

import java.io.IOException;
import java.net.MalformedURLException;
import java.text.*;
import java.util.ArrayList;
import java.util.List;

/**
 * Created by cary on 6/15/17.
 * Updated by RayHu 5/14/18.
 */
public class PdfTableExcel {
	// ExcelObject
	protected ExcelObject excelObject;
	// excel
	protected Excel excel;
	//
	protected boolean setting = false;

	/**
	 * <p>
	 * Description: Constructor
	 * </p>
	 * 
	 * @param excelObject
	 */
	public PdfTableExcel(ExcelObject excelObject) {
		this.excelObject = excelObject;
		this.excel = excelObject.getExcel();
	}

	/**
	 * <p>
	 * Description: 获取转换过的Excel内容Table
	 * </p>
	 * 
	 * @return PdfPTable
	 * @throws BadElementException
	 * @throws MalformedURLException
	 * @throws IOException
	 */
	public PdfPTable getTable() throws BadElementException, MalformedURLException, IOException, DocumentException {
		Sheet sheet = this.excel.getSheet();
		return toParseContent(sheet);
	}

	protected PdfPTable toParseContent(Sheet sheet)
			throws BadElementException, MalformedURLException, IOException, DocumentException {
		int rowlength = sheet.getLastRowNum() + 1;
		List<List<PdfPCell>> cells = new ArrayList<List<PdfPCell>>();
		List<PdfPCell> pdfRow = null;
		List<Float> columnWidths = new ArrayList<Float>();
		for (int i = 0; i < 1000; i++)
			columnWidths.add(0f);
		int columnCount = 0;
		for (int i = 0; i < rowlength; i++) {
			cells.add(pdfRow = new ArrayList<PdfPCell>());
			Row row = sheet.getRow(i);
			int colCount = row.getLastCellNum();
			if (colCount < 0)
				colCount = 0;
			if (colCount > 0) {
				float[] cws = new float[colCount];
				for (int j = 0; j < colCount; j++) {
					if (colCount > columnCount)
						columnCount = colCount;
					Cell cell = row.getCell(j);
					// POI 產生的 Excel 可能會有 null cell
					if (cell == null) {
						cell = row.createCell(j);
					}
					float cw = getPOIColumnWidth(cell);
					cws[cell.getColumnIndex()] = cw;
					columnWidths.set(cell.getColumnIndex(), cw);
					if (isUsed(cell.getColumnIndex(), row.getRowNum())) {
						continue;
					}
					CellRangeAddress range = getColspanRowspanByExcel(row.getRowNum(), cell.getColumnIndex());
					int rowspan = 1;
					int colspan = 1;
					if (range != null) {
						rowspan = range.getLastRow() - range.getFirstRow() + 1;
						colspan = range.getLastColumn() - range.getFirstColumn() + 1;
					}
					// PDF单元格
					PdfPCell pdfpCell = new PdfPCell();

					int[] rgb = getBackgroundColor(cell.getCellStyle());
					pdfpCell.setBackgroundColor(new BaseColor(rgb[0], rgb[1], rgb[2]));
					pdfpCell.setColspan(colspan);
					pdfpCell.setRowspan(rowspan);
					pdfpCell.setVerticalAlignment(getVAlignByExcel(cell.getCellStyle().getVerticalAlignment()));
					pdfpCell.setHorizontalAlignment(getHAlignByExcel(cell.getCellStyle().getAlignment()));
					pdfpCell.setPhrase(getPhrase(cell));
					// pdfpCell.setFixedHeight(this.getPixelHeight(row.getHeightInPoints()));
					pdfpCell.setMinimumHeight(this.getPixelHeight(row.getHeightInPoints()));
					// 不自動換行,應該參數化
					pdfpCell.setNoWrap(true);
					// 中文字才不會黏在邊界線上
					pdfpCell.setPaddingBottom(3);

					addBorderByExcel(pdfpCell, cell.getCellStyle());
					addImageByPOICell(pdfpCell, cell, cw);
					
					pdfRow.add(pdfpCell);
					j += colspan - 1;
				}	
			}
		}

		PdfPTable table = new PdfPTable(columnCount);
		// 設定表格的欄位寬度
		float[] colWidths = new float[columnCount];
		for (int i = 0; i < columnCount; i++)
			colWidths[i] = columnWidths.get(i);
		table.setWidths(colWidths);
		table.setWidthPercentage(100);
		for(List<PdfPCell> row:cells) {
			for (PdfPCell pdfpCell : row) {
				table.addCell(pdfpCell);
			}
			table.completeRow();
		}
		return table;
	}

	protected Phrase getPhrase(Cell cell) {
		HSSFCellStyle cellStyle = (HSSFCellStyle) cell.getCellStyle();
		String formatStr = cellStyle.getDataFormatString();
		// String cellValue = cell.getStringCellValue();
		double cellNumberValue = 0;

		boolean isNumber = false;
		if (cell.getCellType() == Cell.CELL_TYPE_NUMERIC) {
			try {
				cellNumberValue = cell.getNumericCellValue();// Double.parseDouble(cellValue);
				isNumber = true;
			} catch (Exception e) {
			}
		} else if (cell.getCellType() == Cell.CELL_TYPE_FORMULA) {
			FormulaEvaluator evaluator = cell.getSheet().getWorkbook().getCreationHelper().createFormulaEvaluator();
			int resultType = evaluator.evaluateFormulaCell(cell);
			if (resultType == Cell.CELL_TYPE_NUMERIC) {
				try {
					cellNumberValue = cell.getNumericCellValue();
					isNumber = true;
				} catch (Exception e) {
				}
			}
		}

		// 若該 Cell 有特定輸出 Format,則以該 Format 輸出
		if (isNumber) {
			if (!"general".equals(formatStr.toLowerCase())) {
				String numberFormat = formatStr;
				int firstFormatIdx = formatStr.indexOf(";");
				if (firstFormatIdx > 0)
					numberFormat = formatStr.substring(0, firstFormatIdx);
				String formattedValue = new CellNumberFormatter(numberFormat).format(cellNumberValue);
				Phrase phrase = new Phrase(formattedValue, getFontByExcel(cell.getCellStyle()));
				return phrase;
			} else {
				NumberFormat format = new DecimalFormat("#.#");
				Phrase phrase = new Phrase(format.format(cellNumberValue), getFontByExcel(cell.getCellStyle()));
				return phrase;
			}
		} else {
			// 強制轉換為文字類別,如此沒有格式化的數字也能透過 cell.getStringCellValue 取出
			cell.setCellType(Cell.CELL_TYPE_STRING);
			if (this.setting || this.excelObject.getAnchorName() == null) {
				return new Phrase(cell.getStringCellValue(), getFontByExcel(cell.getCellStyle()));
			} else {
				Anchor anchor = new Anchor(cell.getStringCellValue(), getFontByExcel(cell.getCellStyle()));
				anchor.setName(this.excelObject.getAnchorName());
				this.setting = true;
				return anchor;
			}
		}

	}

	public static void main(String[] args) throws Exception {
		NumberFormat format = new DecimalFormat("#.#");
		System.out.println("10.0="+format.format(10.0));
		System.out.println("100.34="+format.format(100.34));
	}

	protected void addImageByPOICell(PdfPCell pdfpCell, Cell cell, float cellWidth)
			throws BadElementException, MalformedURLException, IOException {
		POIImage poiImage = new POIImage().getCellImage(cell);
		byte[] bytes = poiImage.getBytes();
		if (bytes != null) {
			// double cw = cellWidth;
			// double ch = pdfpCell.getFixedHeight();
			//
			// double iw = poiImage.getDimension().getWidth();
			// double ih = poiImage.getDimension().getHeight();
			//
			// double scale = cw / ch;
			//
			// double nw = iw * scale;
			// double nh = ih - (iw - nw);
			//
			// POIUtil.scale(bytes , nw , nh);
			pdfpCell.setVerticalAlignment(Element.ALIGN_MIDDLE);
			pdfpCell.setHorizontalAlignment(Element.ALIGN_CENTER);
			Image image = Image.getInstance(bytes);
			pdfpCell.setImage(image);
		}
	}

	protected float getPixelHeight(float poiHeight) {
		float pixel = poiHeight / 28.6f * 26f;
		return pixel;
	}

	/**
	 * <p>
	 * Description: 此处获取Excel的列宽像素(无法精确实现,期待有能力的朋友进行改善此处)
	 * </p>
	 * 
	 * @param cell
	 * @return 像素宽
	 */
	protected int getPOIColumnWidth(Cell cell) {
		int poiCWidth = excel.getSheet().getColumnWidth(cell.getColumnIndex());
		int colWidthpoi = poiCWidth;
		int widthPixel = 0;
		if (colWidthpoi >= 416) {
			widthPixel = (int) (((colWidthpoi - 416.0) / 256.0) * 8.0 + 13.0 + 0.5);
		} else {
			widthPixel = (int) (colWidthpoi / 416.0 * 13.0 + 0.5);
		}
		return widthPixel;
	}

	protected CellRangeAddress getColspanRowspanByExcel(int rowIndex, int colIndex) {
		CellRangeAddress result = null;
		Sheet sheet = excel.getSheet();
		int num = sheet.getNumMergedRegions();
		for (int i = 0; i < num; i++) {
			CellRangeAddress range = sheet.getMergedRegion(i);
			if (range.getFirstColumn() == colIndex && range.getFirstRow() == rowIndex) {
				result = range;
			}
		}
		return result;
	}

	protected boolean isUsed(int colIndex, int rowIndex) {
		boolean result = false;
		Sheet sheet = excel.getSheet();
		int num = sheet.getNumMergedRegions();
		for (int i = 0; i < num; i++) {
			CellRangeAddress range = sheet.getMergedRegion(i);
			int firstRow = range.getFirstRow();
			int lastRow = range.getLastRow();
			int firstColumn = range.getFirstColumn();
			int lastColumn = range.getLastColumn();
			if (firstRow < rowIndex && lastRow >= rowIndex) {
				if (firstColumn <= colIndex && lastColumn >= colIndex) {
					result = true;
				}
			}
		}
		return result;
	}

	protected Font getFontByExcel(CellStyle style) {
		Workbook wb = excel.getWorkbook();
		// 字体样式索引
		short index = style.getFontIndex();
		org.apache.poi.ss.usermodel.Font font = wb.getFontAt(index);
		// 轉換 POI Font 到 iText Font
		Font itextFont = Resource.getFont((HSSFFont) font);
		Font result = itextFont;
		// 粗體+斜體
		if (font.getBoldweight() == org.apache.poi.ss.usermodel.Font.BOLDWEIGHT_BOLD && font.getItalic()) {
			result.setStyle(Font.BOLDITALIC);
		} else if (font.getBoldweight() == org.apache.poi.ss.usermodel.Font.BOLDWEIGHT_BOLD) { // 粗體
			result.setStyle(Font.BOLD);
		} else if (font.getItalic()) { // 斜體
			result.setStyle(Font.ITALIC);
		}

		// 字体颜色
		int colorIndex = font.getColor();
		HSSFColor color = HSSFColor.getIndexHash().get(colorIndex);
		if (color != null) {
			int rbg = POIUtil.getRGB(color);
			result.setColor(new BaseColor(rbg));
		}

		// 下划线
		FontUnderline underline = FontUnderline.valueOf(font.getUnderline());
		if (underline == FontUnderline.SINGLE) {
			String ulString = Font.FontStyle.UNDERLINE.getValue();
			result.setStyle(ulString);
		}
		return result;
	}

	protected int[] getBackgroundColor(CellStyle style) {
		Color color = style.getFillForegroundColorColor();
		return POIUtil.getColorRGB(color);
	}

	protected int getBackgroundColorByExcel(CellStyle style) {
		Color color = style.getFillForegroundColorColor();
		return POIUtil.getRGB(color);
	}

	/**
	 * TODO:如何將 Excel 中不同樣式粗細的邊線在 PDF 中呈現
	 * 
	 * @param cell
	 * @param style
	 */
	protected void addBorderByExcel(PdfPCell cell, CellStyle style) {
		Workbook wb = excel.getWorkbook();
		cell.setBorderColorLeft(new BaseColor(POIUtil.getBorderRBG(wb, style.getLeftBorderColor())));
		cell.setBorderColorRight(new BaseColor(POIUtil.getBorderRBG(wb, style.getRightBorderColor())));
		cell.setBorderColorTop(new BaseColor(POIUtil.getBorderRBG(wb, style.getTopBorderColor())));
		cell.setBorderColorBottom(new BaseColor(POIUtil.getBorderRBG(wb, style.getBottomBorderColor())));
		// 設定四邊是否有邊線
		short left = getBorderWidth(style.getBorderLeft());
		short right = getBorderWidth(style.getBorderRight());
		short top = getBorderWidth(style.getBorderTop());
		short bottom = getBorderWidth(style.getBorderBottom());
		if (left > 0)
			cell.enableBorderSide(PdfPCell.LEFT);
		else
			cell.disableBorderSide(PdfPCell.LEFT);
		if (right > 0)
			cell.enableBorderSide(PdfPCell.RIGHT);
		else
			cell.disableBorderSide(PdfPCell.RIGHT);
		if (top > 0)
			cell.enableBorderSide(PdfPCell.TOP);
		else
			cell.disableBorderSide(PdfPCell.TOP);
		if (bottom > 0)
			cell.enableBorderSide(PdfPCell.BOTTOM);
		else
			cell.disableBorderSide(PdfPCell.BOTTOM);
	}

	protected short getBorderWidth(short borderType) {
		switch (borderType) {
		case CellStyle.BORDER_DASH_DOT:
		case CellStyle.BORDER_DASH_DOT_DOT:
		case CellStyle.BORDER_DASHED:
		case CellStyle.BORDER_DOTTED:
		case CellStyle.BORDER_HAIR:
		case CellStyle.BORDER_DOUBLE:
		case CellStyle.BORDER_MEDIUM:
		case CellStyle.BORDER_MEDIUM_DASH_DOT:
		case CellStyle.BORDER_MEDIUM_DASH_DOT_DOT:
		case CellStyle.BORDER_MEDIUM_DASHED:
		case CellStyle.BORDER_SLANTED_DASH_DOT:
			return 1;
		case CellStyle.BORDER_NONE:
			return 0;
		case CellStyle.BORDER_THIN:
			return 1;
		case CellStyle.BORDER_THICK:
			return 2;
		default:
			return 1;
		}
	}

	protected int getVAlignByExcel(short align) {
		int result = 0;
		if (align == CellStyle.VERTICAL_BOTTOM) {
			result = Element.ALIGN_BOTTOM;
		}
		if (align == CellStyle.VERTICAL_CENTER) {
			result = Element.ALIGN_MIDDLE;
		}
		if (align == CellStyle.VERTICAL_JUSTIFY) {
			result = Element.ALIGN_JUSTIFIED;
		}
		if (align == CellStyle.VERTICAL_TOP) {
			result = Element.ALIGN_TOP;
		}
		return result;
	}

	protected int getHAlignByExcel(short align) {
		int result = 0;
		if (align == CellStyle.ALIGN_LEFT) {
			result = Element.ALIGN_LEFT;
		}
		if (align == CellStyle.ALIGN_RIGHT) {
			result = Element.ALIGN_RIGHT;
		}
		if (align == CellStyle.ALIGN_JUSTIFY) {
			result = Element.ALIGN_JUSTIFIED;
		}
		if (align == CellStyle.ALIGN_CENTER) {
			result = Element.ALIGN_CENTER;
		}
		return result;
	}
}