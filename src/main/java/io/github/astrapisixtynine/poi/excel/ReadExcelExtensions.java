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
import java.io.IOException;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;


/**
 * The class {@link ReadExcelExtensions} provides methods for read workbook(excel) sheet
 * {@link File} objects
 */
public final class ReadExcelExtensions
{

	/**
	 * Private constructor to prevent instantiation
	 */
	private ReadExcelExtensions()
	{
	}

	/**
	 * Reads the {@link XSSFWorkbook} object from the given {@link File} object
	 *
	 * @param workbookFile
	 *            the workbook(excel) file
	 * @return the {@link XSSFWorkbook} object
	 * @throws IOException
	 *             Signals that an I/O exception has occurred.
	 */
	public static XSSFWorkbook readXSSFWorkbook(File workbookFile) throws IOException
	{
		FileInputStream inputStream = new FileInputStream(workbookFile);
		XSSFWorkbook workbook = new XSSFWorkbook(inputStream);
		return workbook;
	}

	/**
	 * Reads the {@link HSSFWorkbook} object from the given {@link File} object
	 *
	 * @param workbookFile
	 *            the workbook(excel) file
	 * @return the {@link HSSFWorkbook} object
	 * @throws IOException
	 *             Signals that an I/O exception has occurred.
	 */
	public static HSSFWorkbook readHSSFWorkbook(File workbookFile) throws IOException
	{
		FileInputStream inputStream = new FileInputStream(workbookFile);
		HSSFWorkbook workbook = new HSSFWorkbook(inputStream);
		return workbook;
	}
}
