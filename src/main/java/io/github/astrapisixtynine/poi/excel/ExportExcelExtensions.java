/**
 * The MIT License
 *
 * Copyright (C) 2022 Asterios Raptis
 *
 * Permission is hereby granted, free of charge, to any person obtaining
 * a copy of this software and associated documentation files (the
 * "Software"), to deal in the Software without restriction, including
 * without limitation the rights to use, copy, modify, merge, publish,
 * distribute, sublicense, and/or sell copies of the Software, and to
 * permit persons to whom the Software is furnished to do so, subject to
 * the following conditions:
 *
 * The above copyright notice and this permission notice shall be
 * included in all copies or substantial portions of the Software.
 *
 * THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND,
 * EXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF
 * MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND
 * NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE
 * LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER IN AN ACTION
 * OF CONTRACT, TORT OR OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION
 * WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE SOFTWARE.
 */
package io.github.astrapisixtynine.poi.excel;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.List;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.util.NumberToTextConverter;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;


/**
 * The class {@link ExportExcelExtensions} provides methods to export Excel sheets as {@link File}
 * objects and perform various operations on them
 */
public final class ExportExcelExtensions
{

	/**
	 * Private constructor to prevent instantiation
	 */
	private ExportExcelExtensions()
	{
	}

	/**
	 * Exports the given content to an Excel file.
	 *
	 * @param excelFile
	 *            the file to which the content should be written
	 * @param headers
	 *            an array of column headers to be added to the first row of the sheet
	 * @param content
	 *            a two-dimensional array of strings representing the content to be added to the
	 *            sheet
	 * @param sheetName
	 *            the name of the sheet to be created
	 * @throws IOException
	 *             if an I/O error occurs while writing the file
	 */
	public static void exportToExcel(File excelFile, String[] headers, String[][] content,
		final String sheetName) throws IOException
	{
		Workbook workbook = new XSSFWorkbook();
		Sheet sheet = ExcelPoiFactory.newSheet(workbook, sheetName);

		// Create the header row and set style
		Row headerRow = sheet.createRow(0);
		for (int i = 0; i < headers.length; i++)
		{
			Cell cell = headerRow.createCell(i);
			cell.setCellValue(headers[i]);
			cell.setCellStyle(ExcelPoiFactory.newCellStyle(workbook, "Arial", true, (short)12));
		}

		// Add content rows
		for (int i = 0; i < content.length; i++)
		{
			Row row = sheet.createRow(i + 1);
			for (int j = 0; j < content[i].length; j++)
			{
				row.createCell(j).setCellValue(content[i][j]);
			}
		}

		// Auto-size all columns based on the headers length
		for (int i = 0; i < headers.length; i++)
		{
			sheet.autoSizeColumn(i);
		}

		// Write the workbook to the file
		try (FileOutputStream fileOut = new FileOutputStream(excelFile))
		{
			workbook.write(fileOut);
		}
	}

	/**
	 * Checks if the given {@link Row} is empty
	 *
	 * @param row
	 *            the row to check
	 * @return true if the row is empty otherwise false
	 */
	public static boolean isEmpty(Row row)
	{
		if (row == null)
		{
			return true;
		}
		for (Cell cell : row)
		{
			if (!getCellValueAsString(cell).isEmpty())
			{
				return false;
			}
		}
		return true;
	}

	/**
	 * Gets the cell value as an object from the given {@link Cell} object
	 *
	 * @param cell
	 *            the cell
	 * @return the cell value
	 */
	public static Object getCellValue(Cell cell)
	{
		Object result = null;
		if (cell == null)
		{
			return "";
		}
		CellType cellType = cell.getCellType();

		if (CellType.BLANK.equals(cellType))
		{
			result = "";
		}
		else if (CellType.BOOLEAN.equals(cellType))
		{
			result = cell.getBooleanCellValue();
		}
		else if (CellType.ERROR.equals(cellType))
		{
			result = "";
		}
		else if (CellType.FORMULA.equals(cellType))
		{
			result = cell.getCellFormula();
		}
		else if (CellType.NUMERIC.equals(cellType))
		{
			result = cell.getNumericCellValue();
		}
		else if (CellType.STRING.equals(cellType))
		{
			result = cell.getRichStringCellValue().getString();
		}
		return result;
	}

	/**
	 * Exports the given Excel sheet {@link File} and returns a two-dimensional array which holds
	 * the sheets and arrays of the rows
	 *
	 * @param excelSheet
	 *            the Excel sheet {@link File}
	 * @return a two-dimensional array which holds the sheets and arrays of the rows
	 * @throws IOException
	 *             Signals that an I/O exception has occurred
	 * @throws FileNotFoundException
	 *             Signals that the file was not found
	 */
	public static List<String[][]> exportWorkbook(final File excelSheet)
		throws IOException, FileNotFoundException
	{
		final POIFSFileSystem fs = new POIFSFileSystem(new FileInputStream(excelSheet));
		final HSSFWorkbook wb = new HSSFWorkbook(fs);

		final int numberOfSheets = wb.getNumberOfSheets();
		final List<String[][]> sheetList = new ArrayList<>();
		for (int sheetNumber = 0; sheetNumber < numberOfSheets; sheetNumber++)
		{
			HSSFSheet sheet = wb.getSheetAt(sheetNumber);
			final int rows = sheet.getLastRowNum();
			final int columns = sheet.getRow(0).getLastCellNum();
			String[][] excelSheetInTDArray = new String[rows + 1][columns];
			for (int i = 0; i <= rows; i++)
			{
				final HSSFRow row = sheet.getRow(i);
				if (row != null)
				{
					for (int j = 0; j < columns; j++)
					{
						excelSheetInTDArray[i][j] = getCellValueAsString(row.getCell(j));
					}
				}
			}
			sheetList.add(excelSheetInTDArray);
		}
		wb.close();
		return sheetList;
	}

	/**
	 * Gets the cell value as a String from the given {@link Cell} object
	 *
	 * @param cell
	 *            the cell
	 * @return the cell value
	 */
	public static String getCellValueAsString(Cell cell)
	{
		String result = null;
		if (cell == null)
		{
			return "";
		}
		CellType cellType = cell.getCellType();

		if (CellType.BLANK.equals(cellType))
		{
			result = "";
		}
		else if (CellType.BOOLEAN.equals(cellType))
		{
			result = Boolean.toString(cell.getBooleanCellValue());
		}
		else if (CellType.ERROR.equals(cellType))
		{
			result = "";
		}
		else if (CellType.FORMULA.equals(cellType))
		{
			result = cell.getCellFormula();
		}
		else if (CellType.NUMERIC.equals(cellType))
		{
			result = NumberToTextConverter.toText(cell.getNumericCellValue());
		}
		else if (CellType.STRING.equals(cellType))
		{
			result = cell.getRichStringCellValue().getString();
		}
		return result;
	}

	/**
	 * Exports the given Excel sheet {@link File} in a list of lists containing the sheets and lists
	 * of the rows
	 *
	 * @param excelSheet
	 *            the Excel sheet {@link File}
	 * @return a list of lists containing the sheets and lists of the rows
	 * @throws IOException
	 *             Signals that an I/O exception has occurred
	 */
	public static List<List<List<String>>> exportWorkbookAsStringList(final File excelSheet)
		throws IOException
	{
		final HSSFWorkbook wb = ReadExcelExtensions.readHSSFWorkbook(excelSheet);
		return convertToListofLists(wb);
	}

	private static List<List<List<String>>> convertToListofLists(HSSFWorkbook wb) throws IOException
	{
		final int numberOfSheets = wb.getNumberOfSheets();
		final List<List<List<String>>> sl = new ArrayList<>();
		for (int sheetNumber = 0; sheetNumber < numberOfSheets; sheetNumber++)
		{
			HSSFSheet sheet = wb.getSheetAt(sheetNumber);
			final int rows = sheet.getLastRowNum();
			final int columns = sheet.getRow(0).getLastCellNum();
			final List<List<String>> excelSheetList = new ArrayList<>();
			for (int i = 0; i <= rows; i++)
			{
				final HSSFRow row = sheet.getRow(i);
				if (row != null)
				{
					final List<String> reihe = new ArrayList<>();
					for (int j = 0; j < columns; j++)
					{
						reihe.add(getCellValueAsString(row.getCell(j)));
					}
					excelSheetList.add(reihe);
				}
			}
			sl.add(excelSheetList);
		}
		wb.close();
		return sl;
	}

	/**
	 * Replaces null cells with empty cells
	 *
	 * @param excelSheet
	 *            the Excel sheet
	 * @return the HSSF workbook
	 * @throws IOException
	 *             Signals that an I/O exception has occurred
	 * @throws FileNotFoundException
	 *             the file not found exception
	 */
	public static HSSFWorkbook replaceNullCellsIntoEmptyCells(final File excelSheet)
		throws IOException, FileNotFoundException
	{
		final POIFSFileSystem fs = new POIFSFileSystem(new FileInputStream(excelSheet));
		final HSSFWorkbook wb = new HSSFWorkbook(fs);
		final int numberOfSheets = wb.getNumberOfSheets();
		for (int sheetNumber = 0; sheetNumber < numberOfSheets; sheetNumber++)
		{
			HSSFSheet sheet = wb.getSheetAt(sheetNumber);
			final int rows = sheet.getLastRowNum();
			final int columns = sheet.getRow(0).getLastCellNum();
			for (int i = 0; i <= rows; i++)
			{
				final HSSFRow row = sheet.getRow(i);
				if (row != null)
				{
					for (int j = 0; j < columns; j++)
					{
						HSSFCell cell = row.getCell(j);
						if (cell == null)
						{
							cell = row.createCell(j, CellType.BLANK);
						}
					}
				}
			}
		}
		return wb;
	}
}
