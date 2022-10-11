/*
 * Copyright 2015 Viettel ICT. All rights reserved.
 * VIETTEL PROPRIETARY/CONFIDENTIAL. Use is subject to license terms.
 */
package vnm.web.utils;

import java.io.BufferedInputStream;
import java.io.BufferedOutputStream;
import java.io.ByteArrayInputStream;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.io.Writer;
import java.math.BigDecimal;
import java.net.MalformedURLException;
import java.net.URL;
import java.util.ArrayList;
import java.util.Date;
import java.util.HashMap;
import java.util.List;
import java.util.Map;
import java.util.Random;

import javax.servlet.ServletOutputStream;
import javax.servlet.http.HttpServletRequest;
import javax.servlet.http.HttpServletResponse;

import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.streaming.SXSSFSheet;
import org.apache.struts2.ServletActionContext;

import com.viettel.core.exceptions.BusinessException;

import jxl.format.Colour;
import jxl.write.Blank;
import jxl.write.Label;
import jxl.write.Number;
import jxl.write.NumberFormat;
import jxl.write.WritableCell;
import jxl.write.WritableCellFormat;
import jxl.write.WritableFont;
import jxl.write.WriteException;
import net.sf.jasperreports.engine.JRAbstractExporter;
import net.sf.jasperreports.engine.JRDataSource;
import net.sf.jasperreports.engine.JRExporter;
import net.sf.jasperreports.engine.JRExporterParameter;
import net.sf.jasperreports.engine.JRParameter;
import net.sf.jasperreports.engine.JasperFillManager;
import net.sf.jasperreports.engine.JasperPrint;
import net.sf.jasperreports.engine.export.JExcelApiExporter;
import net.sf.jasperreports.engine.export.JRCsvExporter;
import net.sf.jasperreports.engine.export.JRHtmlExporter;
import net.sf.jasperreports.engine.export.JRPdfExporter;
import net.sf.jasperreports.engine.export.JRRtfExporter;
import net.sf.jasperreports.engine.export.JRXlsAbstractExporterParameter;
import net.sf.jasperreports.engine.export.ooxml.JRXlsxExporter;
import net.sf.jxls.transformer.XLSTransformer;
import vnm.web.constant.ConstantManager;
import vnm.web.enumtype.FileExtension;
import vnm.web.enumtype.ShopReportTemplate;
import vnm.web.helper.Configuration;
import vnm.web.helper.R;
import vnm.web.log.LogUtility;

/**
 * Dinh nghia cac ham dung chung cho xuat bao cao
 *
 * @author hunglm16
 * @since 18/09/2015
 */
public class ReportUtils {
	public static String exportFromFormat(FileExtension ext, HashMap<String, Object> parameters, JRDataSource dataSource, ShopReportTemplate shopReport) {
		String outputFile = shopReport.getOutputPath(ext);
		String outputPath = Configuration.getStoreRealPath() + outputFile;
		String outputDownload = Configuration.getExportExcelPath() + outputFile;
		String templatePath = shopReport.getTemplatePath(false, FileExtension.JASPER);
		try {
			JasperPrint jasperPrint = JasperFillManager.fillReport(templatePath, parameters, dataSource);
			//			  String filePath2 = ServletActionContext.getServletContext().getRealPath("report1") + File.separator +"sample_report.html";
			//			   JasperExportManager.exportReportToHtmlFile(jasperPrint, filePath2 );
			JRAbstractExporter exporter = null;
			if (FileExtension.PDF.equals(ext)) {
				exporter = new JRPdfExporter();
			} else if (FileExtension.XLS.equals(ext)) {
				exporter = new JExcelApiExporter();
				//parameters.put(JRXlsAbstractExporterParameter.SHEET_NAMES, "ABC");
				parameters.put(JRParameter.IS_IGNORE_PAGINATION, Boolean.TRUE);
			} else if (FileExtension.XLSX.equals(ext)) {
				HttpServletResponse response = ServletActionContext.getResponse();
				response.setHeader("Content-disposition", "attachment");
				response.setContentType("application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
				response.setCharacterEncoding(ConstantManager.UTF_8);
				exporter = new JRXlsxExporter();
				//parameters.put(JRXlsAbstractExporterParameter.SHEET_NAMES, "ABC");
				parameters.put(JRParameter.IS_IGNORE_PAGINATION, Boolean.TRUE);
			} else if (FileExtension.DOC.equals(ext)) {
				exporter = new JRRtfExporter();
			} else if (FileExtension.CSV.equals(ext)) {
				HttpServletResponse response = ServletActionContext.getResponse();
				response.setHeader("Content-disposition", "attachment");
				response.setContentType("application/octet-stream");
				response.setCharacterEncoding(ConstantManager.UTF_8);
				exporter = new JRCsvExporter();
				parameters.put(JRParameter.IS_IGNORE_PAGINATION, Boolean.TRUE);
			}
			if (exporter != null) {
				exporter.setParameter(JRExporterParameter.JASPER_PRINT, jasperPrint);
				exporter.setParameter(JRExporterParameter.OUTPUT_FILE_NAME, outputPath);
				exporter.exportReport();
			}
		} catch (Exception e) {
			e.printStackTrace();
			LogUtility.logError(e, e.getMessage(), ReportUtils.class, null);
		}
		return outputDownload;
	}

	/**
	 * Dung cho chuc nang Zip File
	 *
	 * @author hunglm16
	 * @since MAY 26,2014
	 */
	public static String exportFromFormatNew(FileExtension ext, HashMap<String, Object> parameters, JRDataSource dataSource, ShopReportTemplate shopReport) {
		String outputFile = shopReport.getOutputPath(ext);
		String outputPath = Configuration.getStoreRealPath() + outputFile;
		String templatePath = shopReport.getTemplatePath(false, FileExtension.JASPER);
		try {
			JasperPrint jasperPrint = JasperFillManager.fillReport(templatePath, parameters, dataSource);
			JRAbstractExporter exporter = null;
			if (FileExtension.PDF.equals(ext)) {
				exporter = new JRPdfExporter();
			} else if (FileExtension.XLS.equals(ext)) {
				exporter = new JExcelApiExporter();
				//parameters.put(JRXlsAbstractExporterParameter.SHEET_NAMES, "ABC");
				parameters.put(JRParameter.IS_IGNORE_PAGINATION, Boolean.TRUE);
			} else if (FileExtension.XLSX.equals(ext)) {
				HttpServletResponse response = ServletActionContext.getResponse();
				response.setHeader("Content-disposition", "attachment");
				response.setContentType("application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
				response.setCharacterEncoding(ConstantManager.UTF_8);
				exporter = new JRXlsxExporter();
				//parameters.put(JRXlsAbstractExporterParameter.SHEET_NAMES, "ABC");
				parameters.put(JRParameter.IS_IGNORE_PAGINATION, Boolean.TRUE);
			} else if (FileExtension.DOC.equals(ext)) {
				exporter = new JRRtfExporter();
			} else if (FileExtension.CSV.equals(ext)) {
				HttpServletResponse response = ServletActionContext.getResponse();
				response.setHeader("Content-disposition", "attachment");
				response.setContentType("application/octet-stream");
				response.setCharacterEncoding(ConstantManager.UTF_8);
				exporter = new JRCsvExporter();
				parameters.put(JRParameter.IS_IGNORE_PAGINATION, Boolean.TRUE);
			}
			if (exporter != null) {
				exporter.setParameter(JRExporterParameter.JASPER_PRINT, jasperPrint);
				exporter.setParameter(JRExporterParameter.OUTPUT_FILE_NAME, outputPath);
				exporter.exportReport();
			}
		} catch (Exception e) {
			LogUtility.logError(e, e.getMessage(), ReportUtils.class, null);
		}
		return outputPath;
	}

	/**
	 * Them tham so sheetName de dat ten sheet khi xuat file excel.
	 *
	 * @author hunglm16
	 * @param ext
	 * @param parameters
	 * @param dataSource
	 * @param shopReport
	 * @param sheetName
	 * @return
	 * @since 18/09/2015
	 */
	public static String exportFromFormat(FileExtension ext, HashMap<String, Object> parameters, JRDataSource dataSource, ShopReportTemplate shopReport, String sheetName) {
		String outputFile = shopReport.getOutputPath(ext);
		String outputPath = Configuration.getStoreRealPath() + outputFile;
		String outputDownload = Configuration.getExportExcelPath() + outputFile;
		String templatePath = shopReport.getTemplatePath(false, FileExtension.JASPER);
		try {
			JasperPrint jasperPrint = JasperFillManager.fillReport(templatePath, parameters, dataSource);
			JRAbstractExporter exporter = null;
			if (FileExtension.PDF.equals(ext)) {
				exporter = new JRPdfExporter();
			} else if (FileExtension.XLS.equals(ext)) {
				exporter = new JExcelApiExporter();
				//parameters.put(JRXlsAbstractExporterParameter.SHEET_NAMES, "ABC");
				parameters.put(JRParameter.IS_IGNORE_PAGINATION, Boolean.TRUE);
			} else if (FileExtension.DOC.equals(ext)) {
				exporter = new JRRtfExporter();
			} else if (FileExtension.CSV.equals(ext)) {
				HttpServletResponse response = ServletActionContext.getResponse();
				response.setHeader("Content-disposition", "attachment");
				response.setContentType("application/octet-stream");
				response.setCharacterEncoding(ConstantManager.UTF_8);
				exporter = new JRCsvExporter();
				parameters.put(JRParameter.IS_IGNORE_PAGINATION, Boolean.TRUE);
			}
			if (exporter != null) {
				exporter.setParameter(JRExporterParameter.JASPER_PRINT, jasperPrint);
				exporter.setParameter(JRExporterParameter.OUTPUT_FILE_NAME, outputPath);
				if (!StringUtil.isNullOrEmpty(sheetName)) {
					exporter.setParameter(JRXlsAbstractExporterParameter.SHEET_NAMES, new String[] { sheetName });
				}
				exporter.exportReport();
			}
		} catch (Exception e) {
			LogUtility.logError(e, e.getMessage(), ReportUtils.class, null);
		}
		return outputDownload;
	}

	/**
	 * Gets the max number row
	 *
	 * @author hunglm16
	 * @param numCol
	 * @return
	 * @since 18/09/2015
	 */
	public static int getMaxNumberRow(int numCol) {
		return (int) (Configuration.getMaxNumberCellReport() / numCol);
	}

	public static int getMaxNumRowCrosstabs(int sizeListAuto) {
		return (int) (Configuration.getMaxNumberCellReport() / sizeListAuto);
	}

	/**
	 * Gets the max report in page.
	 *
	 * @return the max report in page
	 */
	public static int getMaxReportInPage() {
		return 50;
	}

	/**
	 * Export HTML
	 *
	 * @author hunglm16
	 * @param response
	 * @param parameters
	 * @param dataSource
	 * @param shopReport
	 * @since 18/09/2015
	 */
	public static void exportHtml(HttpServletResponse response, HashMap<String, Object> parameters, JRDataSource dataSource, ShopReportTemplate shopReport) {
		String templatePath = shopReport.getTemplatePath(false, FileExtension.JASPER);
		ServletOutputStream ouputStream = null;
		try {
			parameters.put(JRParameter.IS_IGNORE_PAGINATION, Boolean.TRUE);
			JasperPrint jasperPrint = JasperFillManager.fillReport(templatePath, parameters, dataSource);
			ouputStream = response.getOutputStream();
			JRExporter exporter = null;
			exporter = new JRHtmlExporter();
			exporter.setParameter(JRExporterParameter.OUTPUT_STREAM, false);
			exporter.setParameter(JRExporterParameter.JASPER_PRINT, jasperPrint);
			exporter.setParameter(JRExporterParameter.OUTPUT_STREAM, ouputStream);
			exporter.exportReport();
		} catch (Exception e) {
			LogUtility.logError(e, e.getMessage(), ReportUtils.class, null);
			System.out.print(e.getMessage());
		} finally {
			if (ouputStream != null) {
				try {
					ouputStream.close();
				} catch (IOException e) {
					LogUtility.logError(e, e.getMessage(), ReportUtils.class, null);
				}
			}
		}
	}

	public static void exportCsvFromFormat(FileExtension ext, HashMap<String, Object> parameters, JRDataSource dataSource, ShopReportTemplate shopReport) {
		String outputFile = shopReport.getOutputPath(ext);
		String outputPath = Configuration.getStoreRealPath() + outputFile;
		//String outputDownload = Configuration.getExportExcelPath() + outputFile;
		String templatePath = shopReport.getTemplatePath(false, FileExtension.JASPER);
		try {
			JasperPrint jasperPrint = JasperFillManager.fillReport(templatePath, parameters, dataSource);
			HttpServletResponse response = ServletActionContext.getResponse();
			JRCsvExporter csvExporter = new JRCsvExporter();
			parameters.put(JRParameter.IS_IGNORE_PAGINATION, Boolean.TRUE);
			csvExporter.setParameter(JRExporterParameter.JASPER_PRINT, jasperPrint);

			StringBuffer buffer = new StringBuffer();
			csvExporter.setParameter(JRExporterParameter.OUTPUT_STRING_BUFFER, buffer);
			csvExporter.setParameter(JRExporterParameter.OUTPUT_FILE_NAME, outputPath);
			csvExporter.exportReport();

			response.setContentType("application/octet-stream");
			response.setHeader("Content-Disposition", "attachment;filename=" + outputFile);
			response.setCharacterEncoding(ConstantManager.UTF_8_ENCODING);
			response.getOutputStream().write(buffer.toString().getBytes());

			StringBuffer sb = new StringBuffer();
			InputStream inputStream = new ByteArrayInputStream(sb.toString().getBytes("UTF-8"));
			ServletOutputStream out = response.getOutputStream();

			byte[] outputByte = new byte[4096];
			//copy binary contect to output stream
			while (inputStream.read(outputByte, 0, 4096) != -1) {
				out.write(outputByte, 0, 4096);
			}
			inputStream.close();
			//byte-order marker (BOM)
			byte b[] = { (byte) 0xEF, (byte) 0xBB, (byte) 0xBF };
			//insert BOM byte array into outputStream
			out.write(b);
			out.flush();
			out.close();
		} catch (Exception e) {
			LogUtility.logError(e, e.getMessage(), ReportUtils.class, null);
		}
	}

	/**
	 * Gets the vinamilk logo.
	 *
	 * @return the vinamilk logo
	 */
	public static InputStream getVinamilkLogoStream() {
		return ReportUtils.class.getClassLoader().getResourceAsStream("vinamilk80.jpg");
	}

	/**
	 * Gets the vinamilk logo real path.
	 *
	 * @param request
	 *            the request
	 * @return the vinamilk logo real path
	 */
	public static String getVinamilkLogoRealPath(HttpServletRequest request) {
		return request.getServletContext().getRealPath("resources/images/vinamilk80.jpg");
	}

	/**
	 * Gets the vinamilk logo real path .png.
	 *
	 * @param request
	 *            the request
	 * @return the vinamilk logo real path .png
	 */
	public static String getVinamilkLogoRealPathPNG(HttpServletRequest request) {
		return request.getServletContext().getRealPath("resources/images/vinamilk90.png");
	}

	/**
	 * Gets the vinamilk logo real path.
	 *
	 * @param request
	 *            the request
	 * @return the icon uncheck
	 */
	public static String getVinamilkUnchecked(HttpServletRequest request) {
		return request.getServletContext().getRealPath("resources/images/unchecked.png");
	}

	/**
	 * Gets the vinamilk logo real path.
	 *
	 * @param request
	 *            the request
	 * @return the icon check
	 */
	public static String getVinamilkChecked(HttpServletRequest request) {
		return request.getServletContext().getRealPath("resources/images/checked.png");
	}

	/**
	 * Gets the vinamilk logo real path.
	 *
	 * @author hunglm16
	 * @param request
	 *            the request
	 * @return the icon check
	 */
	public static String getVinamilkCheckedSmall(HttpServletRequest request) {
		return request.getServletContext().getRealPath("resources/images/checked_small.png");
	}

	/**
	 * Gets the vinamilk logo real path.
	 *
	 * @author hunglm16
	 * @param request
	 *            the request
	 * @return the icon uncheck
	 */
	public static String getVinamilkUncheckedSmall(HttpServletRequest request) {
		return request.getServletContext().getRealPath("resources/images/unChecked_small.png");
	}

	/**
	 * Export excel jxls.
	 *
	 * @param beans
	 *            the beans
	 * @param tabletPortalReport
	 *            the tablet portal report
	 * @return the string
	 * @author thongnm
	 * @throws IOException
	 * @since Apr 8, 2013
	 */
	public static String exportExcelJxls(HashMap<String, Object> beans, ShopReportTemplate tabletPortalReport, FileExtension ext) throws BusinessException, IOException {
		String outputFile = tabletPortalReport.getOutputPath(ext);
		String outputPath = Configuration.getStoreRealPath() + outputFile;
		String outputDownload = Configuration.getExportExcelPath() + outputFile;
		String templatePath = tabletPortalReport.getTemplatePath(false, FileExtension.XLS);
		InputStream inputStream = null;
		OutputStream os = null;
		Workbook resultWorkbook = null;
		try {
			//Configuration.getStoreRealPath();
			inputStream = new BufferedInputStream(new FileInputStream(templatePath));
			XLSTransformer transformer = new XLSTransformer();
			resultWorkbook = transformer.transformXLS(inputStream, beans);
			os = new BufferedOutputStream(new FileOutputStream(outputPath));
			resultWorkbook.write(os);
			os.flush();
			//os.close();

		} catch (Exception e) {
			throw new BusinessException(e);
		} finally {
			if (inputStream != null) {
				try {
					inputStream.close();
				} catch (IOException e) {
					LogUtility.logErrorStandard(e, e.getMessage(), null);
				}
			}
			if (os != null) {
				try {
					os.close();
				} catch (IOException e) {
					LogUtility.logErrorStandard(e, e.getMessage(), null);
				}
			}
			if (resultWorkbook != null) {
				try {
					resultWorkbook.close();
				} catch (IOException e) {
					LogUtility.logErrorStandard(e, e.getMessage(), null);
				}
			}
		}
		return outputDownload;
	}

	/**
	 * @author sangtn
	 * @param c
	 * @param r
	 * @param val
	 * @param isBold
	 * @param alignment
	 * @param backgroundColour
	 * @param numberFormat
	 * @return
	 */
	public static WritableCell addCell(int c, int r, Object val, boolean isBold, jxl.format.Alignment alignment, Colour backgroundColour, NumberFormat numberFormat) {
		try {
			WritableFont cellFont = new WritableFont(WritableFont.TIMES, 10);
			WritableFont cellFontBold = new WritableFont(WritableFont.TIMES, 10, WritableFont.BOLD);
			WritableCellFormat cell = new WritableCellFormat(cellFont);
			WritableCellFormat cellNumber = new WritableCellFormat(cellFont, numberFormat);

			cell.setBorder(jxl.format.Border.ALL, jxl.format.BorderLineStyle.THIN);
			cell.setWrap(true);
			cell.setVerticalAlignment(jxl.format.VerticalAlignment.CENTRE);

			cellNumber.setBorder(jxl.format.Border.ALL, jxl.format.BorderLineStyle.THIN);
			cellNumber.setVerticalAlignment(jxl.format.VerticalAlignment.CENTRE);

			if (isBold == true) {
				cell.setFont(cellFontBold);
				cellNumber.setFont(cellFontBold);
			}
			if (alignment != null) {
				cell.setAlignment(alignment);
				cellNumber.setAlignment(alignment);
			}
			if (backgroundColour != null) {
				cell.setBackground(backgroundColour);
				cellNumber.setBackground(backgroundColour);
			}

			if (val == null) {
				return new Blank(c, r, cell);
			} else if (val instanceof BigDecimal) {
				return new Number(c, r, ((BigDecimal) val).doubleValue(), cellNumber);
			} else if (val instanceof Integer) {
				return new Number(c, r, (Integer) val, cell);
			} else if (val instanceof String) {
				return new Label(c, r, (String) val, cell);
			} else if (val instanceof Date) {
				return new Label(c, r, DateUtil.toDateString((Date) val), cell);
			}
		} catch (WriteException e) {
			LogUtility.logError(e, e.getMessage(), ReportUtils.class, null);
		}
		return null;

	}

	public static WritableCell addCell(int c, int r, Object val, boolean isBold, int sizeFont, jxl.format.Alignment alignment, Colour backgroundColour, NumberFormat numberFormat) {
		try {
			WritableFont cellFont = new WritableFont(WritableFont.TIMES, sizeFont);
			WritableFont cellFontBold = new WritableFont(WritableFont.TIMES, sizeFont, WritableFont.BOLD);
			WritableCellFormat cell;
			if (numberFormat != null) {
				cell = new WritableCellFormat(cellFont, numberFormat);
			} else {
				cell = new WritableCellFormat(cellFont);
			}

			cell.setBorder(jxl.format.Border.ALL, jxl.format.BorderLineStyle.THIN);
			cell.setVerticalAlignment(jxl.format.VerticalAlignment.CENTRE);

			if (isBold == true) {
				cell.setFont(cellFontBold);
			}
			if (alignment != null) {
				cell.setAlignment(alignment);
			}
			if (backgroundColour != null) {
				cell.setBackground(backgroundColour);
			}

			if (val == null) {
				return new Blank(c, r, cell);
			} else if (val instanceof BigDecimal) {
				return new Number(c, r, ((BigDecimal) val).doubleValue(), cell);
			} else if (val instanceof Double) {
				return new Number(c, r, (Double) val, cell);
			} else if (val instanceof Float) {
				return new Number(c, r, (Float) val, cell);
			} else if (val instanceof Integer) {
				return new Number(c, r, (Integer) val, cell);
			} else if (val instanceof String) {
				return new Label(c, r, (String) val, cell);
			} else if (val instanceof Date) {
				return new Label(c, r, DateUtil.toDateString((Date) val), cell);
			}
		} catch (WriteException e) {
			LogUtility.logError(e, e.getMessage(), ReportUtils.class, null);
		}
		return null;
	}

	/**
	 * Set RowHeight by height for row in index position (by point) Dat do cao
	 * height cho dong index (by point)
	 */
	public void setRowHeight(SXSSFSheet sheet, int rowIndex, int height) {
		Row row = sheet.getRow(rowIndex) == null ? sheet.createRow(rowIndex) : sheet.getRow(rowIndex);
		row.setHeight((short) (height * 20)); // 1/20 of a point
	}

	/**
	 * Set ColumnWidth by width for column in index position (by pixel) Dat do
	 * rong width cho cot index (by pixel)
	 */
	public void setColumnWidth(SXSSFSheet sheet, int colIndex, int width) {
		sheet.setColumnWidth(colIndex, (int) ((width * 256) / (ConstantManager.XSSF_MAX_DIGIT_WIDTH + ConstantManager.XSSF_CHARACTER_DEFAULT_PADDING)));
	}

	/**
	 * Set ColumnWidthfor multiple Column, start from startIndex (by pixel) Dat
	 * do rong cho nhieu cot, bat dau tu cot thu startIndex (by pixel)
	 */
	public void setColumnsWidth(SXSSFSheet sheet, Integer startIndex, Integer... widths) {
		for (int i = 0, sizeTmp = widths.length; i < sizeTmp; i++) {
			sheet.setColumnWidth(i + startIndex, (int) ((widths[i] * 256) / (ConstantManager.XSSF_MAX_DIGIT_WIDTH + ConstantManager.XSSF_CHARACTER_DEFAULT_PADDING)));
		}
	}

	/**
	 * Set RowHeight for multiple row, start from startIndex (by point) Dat do
	 * cao cho nhieu dong, bat dau tu cot thu startIndex (by point)
	 */
	public void setRowsHeight(SXSSFSheet sheet, Integer startIndex, Integer... heights) {
		for (int i = 0, sizeTmp = heights.length; i < sizeTmp; i++) {
			this.setRowHeight(sheet, i + startIndex, heights[i]);
		}
	}

	public void flushAllRowAndGroupLines(SXSSFSheet sheet, List<Integer> lstStartIndex, List<Integer> lstEndIndex, int curEndIndex, int rowNumInWindow) throws IOException {
		int curIndex = (curEndIndex - rowNumInWindow) + 1;
		List<Integer> lstCurStartIndex = new ArrayList<>();
		List<Integer> lstCurEndIndex = new ArrayList<>();
		lstCurStartIndex.addAll(lstStartIndex);
		lstCurEndIndex.addAll(lstEndIndex);

		if (((lstStartIndex != null) && (lstStartIndex.size() > 0)) && ((lstEndIndex != null) && (lstEndIndex.size() > 0))) {
			for (; curIndex <= curEndIndex; curIndex++) {
				for (int i = 0, sizeStart = lstStartIndex.size(); i < sizeStart; i++) {
					if ((curIndex >= lstCurStartIndex.get(i)) && (curIndex < lstCurEndIndex.get(i))) {
						sheet.groupRow(curIndex, curIndex);
					}
				}
			}
			sheet.flushRows();
			return;
		} else {
			sheet.flushRows();
			return;
		}
	}

	public void groupRows(SXSSFSheet sheet, int startIndex, int endIndex, Integer rowNumInWindow, List<Integer> lstStartIndex, List<Integer> lstEndIndex) throws IOException {

		if (((lstStartIndex != null) && (lstStartIndex.size() > 0)) && ((lstEndIndex != null) && (lstEndIndex.size() > 0))) {
			if (((rowNumInWindow != null) && (rowNumInWindow >= 0))) {
				if ((endIndex - rowNumInWindow) > startIndex) {
					return;
				}
			}
			for (int i = 0; i < lstStartIndex.size(); i++) {
				if ((startIndex == lstStartIndex.get(i)) && (endIndex == lstEndIndex.get(i))) {
					lstStartIndex.remove(i);
					lstEndIndex.remove(i);
				}
			}
			sheet.groupRow(startIndex, endIndex);
		} else {
			sheet.groupRow(startIndex, endIndex);
		}
		return;
	}

	/**
	 * Xuat file excel bang template su dung jxls
	 *
	 * @param beans
	 * @param templatePath
	 *            - duong dan den thu muc chua file template
	 * @param templateName
	 *            - ten file template (khi xuat ra se gan them thoi gian vao
	 *            phia sau + .xls)
	 * @param templateNameRes
	 *            - chuoi key trong file resource, neu templateName null thi su
	 *            dung de lay ten file bao cao trong resource
	 *
	 * @author lacnv1
	 * @since Apr 08, 2015
	 */
	public static String exportXLSWithJxls(Map<String, Object> beans, String templatePath, String templateName, String templateNameRes) throws Exception {
		InputStream inputStream = null;
		OutputStream os = null;
		try {
			if (beans == null) {
				throw new IllegalArgumentException("beans: null");
			}
			if (StringUtil.isNullOrEmpty(templatePath)) {
				throw new IllegalArgumentException("templatePath: null");
			}
			if (StringUtil.isNullOrEmpty(templateName) && StringUtil.isNullOrEmpty(templateNameRes)) {
				throw new IllegalArgumentException("fileName + fileNameRes: null");
			}
			if (StringUtil.isNullOrEmpty(templateName)) {
				templateName = R.getResource(templateNameRes);
			}

			String folder = ServletActionContext.getServletContext().getRealPath("/") + templatePath;
			StringBuilder sb = new StringBuilder(folder).append(templateName);
			String sTmp = templateName.toUpperCase();
			if (!sTmp.endsWith(FileExtension.XLS.getValue().toUpperCase())) {
				sb.append(FileExtension.XLS.getValue());
			}
			String templateFileName = sb.toString();
			templateFileName = templateFileName.replace('/', File.separatorChar);

			sb = new StringBuilder(Configuration.getResourceString(ConstantManager.VI_LANGUAGE, templateName)).append(DateUtil.toDateString(DateUtil.now(), DateUtil.DATE_FORMAT_EXCEL_FILE)).append(FileExtension.XLS.getValue());
			String outputName = sb.toString();
			sb = null;
			String exportFileName = Configuration.getStoreRealPath() + outputName;

			inputStream = new BufferedInputStream(new FileInputStream(templateFileName));
			XLSTransformer transformer = new XLSTransformer();
			Workbook resultWorkbook = transformer.transformXLS(inputStream, beans);
			// inputStream.close();
			os = new BufferedOutputStream(new FileOutputStream(exportFileName));
			resultWorkbook.write(os);
			os.flush();
			// os.close();
			return Configuration.getExportExcelPath() + outputName;
		} catch (Exception ex) {
			throw ex;
		} finally {
			if (inputStream != null) {
				inputStream.close();
			}
			if (os != null) {
				os.close();
			}
		}
	}

	/** End SXSSF range */

	//public static char csvSep = '\t';
	public static final char csvSep = ',';

	public static void appendCSVsepPrefix(Writer writer, Object val) throws Exception {
		if (val == null) {
			writer.append(ReportUtils.csvSep);
		} else {
			writer.append(ReportUtils.csvSep).append("\"" + val.toString() + "\"");
		}
	}

	public static void appendCSVsepSuffix(Writer writer, Object val) throws Exception {
		if (val == null) {
			writer.append(ReportUtils.csvSep);
		} else {
			writer.append("\"" + val.toString() + "\"").append(ReportUtils.csvSep);
		}
	}

	public static void appendStringCSVsepPrefix(Writer writer, String text) throws Exception {
		if (text == null) {
			writer.append(ReportUtils.csvSep);
		} else {
			writer.append(ReportUtils.csvSep).append("\"" + text + "\"");
		}
	}

	public static void appendStringCSVsepSuffix(Writer writer, String text) throws Exception {
		if (text == null) {
			writer.append(ReportUtils.csvSep);
		} else {
			writer.append("\"" + text + "\"").append(ReportUtils.csvSep);
		}
	}

	/**
	 * Lay ten xuat file theo chuan yeu cau
	 *
	 * @author hunglm16
	 * @param name
	 * @param fDate
	 * @param tDate
	 * @param fileExtension
	 * @return
	 * @since 18/09/2015
	 */
	public String getReportNameFormat(String name, String fDate, String tDate, String fileExtension) {
		String nameRpt = "Viettel_";
		if (fDate.isEmpty() && tDate.isEmpty()) {
			nameRpt = nameRpt + name;
		} else {
			Date fD = DateUtil.parse(fDate, DateUtil.DATE_FORMAT_DDMMYYYY);
			Date tD = DateUtil.parse(tDate, DateUtil.DATE_FORMAT_DDMMYYYY);
			fDate = DateUtil.toDateString(fD, DateUtil.DATE_FORMAT_CSV);
			tDate = DateUtil.toDateString(tD, DateUtil.DATE_FORMAT_CSV);
			nameRpt = nameRpt + name + "_" + fDate + "-" + tDate;
		}
		Random rand = new Random();
		Integer n = rand.nextInt(100) + 1;
		return nameRpt + "_" + n.toString() + fileExtension;
	}

	/**
	 * Export excel jxls.
	 *
	 * @param beans
	 *            the beans
	 * @param tabletPortalReport
	 *            the tablet portal report
	 * @return the string
	 * @author vuongmq
	 * @throws IOException
	 * @date Mar 20, 2015
	 */
	public static String exportExcelJxlsx(HashMap<String, Object> beans, ShopReportTemplate tabletPortalReport, FileExtension ext) throws BusinessException, IOException {
		String outputFile = tabletPortalReport.getOutputPath(ext);
		String outputPath = Configuration.getStoreRealPath() + outputFile;
		String outputDownload = Configuration.getExportExcelPath() + outputFile;
		String templatePath = tabletPortalReport.getTemplatePathExcel(true, ext);
		templatePath = templatePath.replace('/', File.separatorChar);
		InputStream inputStream = null;
		FileInputStream fileInputStream = null;
		FileOutputStream fileOutStream = null;
		OutputStream os = null;
		Workbook resultWorkbook = null;
		try {
			//Configuration.getStoreRealPath();
			fileInputStream = new FileInputStream(templatePath);
			inputStream = new BufferedInputStream(fileInputStream);
			XLSTransformer transformer = new XLSTransformer();
			resultWorkbook = transformer.transformXLS(inputStream, beans);
			fileOutStream = new FileOutputStream(outputPath);
			os = new BufferedOutputStream(fileOutStream);
			resultWorkbook.write(os);
			os.flush();
		} catch (Exception e) {
			throw new BusinessException(e);
		} finally {
			if (inputStream != null) {
				try {
					inputStream.close();
				} catch (IOException e) {
					LogUtility.logErrorStandard(e, e.getMessage(), null);
				}
			}
			if (fileInputStream != null) {
				try {
					fileInputStream.close();
				} catch (IOException e) {
					LogUtility.logErrorStandard(e, e.getMessage(), null);
				}
			}
			if (fileOutStream != null) {
				try {
					fileOutStream.close();
				} catch (IOException e) {
					LogUtility.logErrorStandard(e, e.getMessage(), null);
				}
			}
			if (os != null) {
				try {
					os.close();
				} catch (IOException e) {
					LogUtility.logErrorStandard(e, e.getMessage(), null);
				}
			}
			if (resultWorkbook != null) {
				try {
					resultWorkbook.close();
				} catch (IOException e) {
					LogUtility.logErrorStandard(e, e.getMessage(), null);
				}
			}
		}
		return outputDownload;
	}

	/**
	 * Bo sung tham so report token vao duong dan download file
	 *
	 * @author tuannd20
	 * @param downloadUrl
	 *            Duong dan goc
	 * @param reportToken
	 *            Report token
	 * @return Duong dan co tham so report token
	 * @since 05/05/2015
	 */
	public static String addReportTokenToDownloadFileUrl(String downloadUrl, String reportToken) {
		final String QUERY_HEADER = "?v=";
		final String QUERY_PARAM = "&v=";
		if (!StringUtil.isNullOrEmpty(reportToken)) {
			try {
				URL url = new java.net.URL(downloadUrl);
				String query = url.getQuery();
				if (StringUtil.isNullOrEmpty(query)) {
					downloadUrl += QUERY_HEADER;
				} else {
					downloadUrl += QUERY_PARAM;
				}
				downloadUrl += reportToken;
			} catch (MalformedURLException e) {
				LogUtility.logError(e, e.getMessage(), ReportUtils.class, null);
			}
		}
		return downloadUrl;
	}
}
