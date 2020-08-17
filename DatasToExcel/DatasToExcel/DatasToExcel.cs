using System;
using System.IO;
using System.Linq;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace DatasToExcel
{
    public static class DatasToExcel
    {
        /// <summary>
        /// Generate Excel file from 2D array.
        /// </summary>
        /// <typeparam name="T">The type of values in array.</typeparam>
        /// <param name="data">The 2D array, first dimension is row and second dimension is column.</param>
        public static MemoryStream GenerateExcel<T>(this T[,] data)
        {
            MemoryStream stream = new MemoryStream();
            GenerateExcel(data, stream, false);
            return stream;
        }

        /// <summary>
        /// Generate Excel file from 2D array.
        /// </summary>
        /// <typeparam name="T">The type of values in array.</typeparam>
        /// <param name="data">The 2D array, first dimension is row and second dimension is column.</param>
        /// <param name="headerFirstRow">Is the first row in data is header or not.</param>
        public static MemoryStream GenerateExcel<T>(this T[,] data, bool headerFirstRow)
        {
            MemoryStream stream = new MemoryStream();
            GenerateExcel(data, stream, headerFirstRow);
            return stream;
        }

        /// <summary>
        /// Generate Excel file from 2D array.
        /// </summary>
        /// <typeparam name="T">The type of values in array.</typeparam>
        /// <param name="data">The 2D array, first dimension is row and second dimension is column.</param>
        /// <param name="filename">The output path to save the Excel file.</param>
        public static void GenerateExcel<T>(this T[,] data, string filename)
        {
            GenerateExcel(data, new FileStream(filename, FileMode.Create, FileAccess.ReadWrite), false);
        }

        /// <summary>
        /// Generate Excel file from 2D array.
        /// </summary>
        /// <typeparam name="T">The type of values in array.</typeparam>
        /// <param name="data">The 2D array, first dimension is row and second dimension is column.</param>
        /// <param name="filename">The output path to save the Excel file.</param>
        /// <param name="headerFirstRow">Is the first row in data is header or not.</param>
        public static void GenerateExcel<T>(this T[,] data, string filename, bool headerFirstRow)
        {
            GenerateExcel(data, new FileStream(filename, FileMode.Create, FileAccess.ReadWrite), headerFirstRow);
        }

        /// <summary>
        /// Generate Excel file from 2D array.
        /// </summary>
        /// <typeparam name="T">The type of values in array.</typeparam>
        /// <param name="data">The 2D array, first dimension is row and second dimension is column.</param>
        /// <param name="stream">The output stream to save the Excel file.</param>
        public static void GenerateExcel<T>(this T[,] data, Stream stream)
        {
            GenerateExcel(data, stream, false);
        }

        /// <summary>
        /// Generate Excel file from 2D array.
        /// </summary>
        /// <typeparam name="T">The type of values in array.</typeparam>
        /// <param name="data">The 2D array, first dimension is row and second dimension is column.</param>
        /// <param name="stream">The output stream to save the Excel file.</param>
        /// <param name="headerFirstRow">Is the first row in data is header or not.</param>
        public static void GenerateExcel<T>(this T[,] data, Stream stream, bool headerFirstRow)
        {
            try
            {
                using (SpreadsheetDocument doc = SpreadsheetDocument.Create(stream, DocumentFormat.OpenXml.SpreadsheetDocumentType.Workbook, true))
                {
                    WorkbookPart workbook = doc.WorkbookPart ?? doc.AddWorkbookPart();
                    workbook.Workbook ??= new Workbook();

                    WorksheetPart worksheetPart = workbook.AddNewPart<WorksheetPart>();
                    worksheetPart.Worksheet ??= new Worksheet();

                    if (headerFirstRow == true)
                    {
                        SheetViews sheetViews = new SheetViews();
                        SheetView sheetView = new SheetView() { TabSelected = true, WorkbookViewId = 0 };
                        Pane pane = new Pane() { ActivePane = PaneValues.BottomLeft, State = PaneStateValues.Frozen, TopLeftCell = "A2", VerticalSplit = 1D };
                        Selection selection = new Selection() { Pane = PaneValues.BottomLeft };

                        sheetView.Append(pane);
                        sheetView.Append(selection);
                        sheetViews.Append(sheetView);

                        worksheetPart.Worksheet.Append(sheetViews);
                    }

                    worksheetPart.Worksheet.Save();

                    Columns columns = new Columns();

                    for (int i = 1; i <= data.GetLength(1); i++)
                    {
                        columns.AppendChild(new Column() { Min = (uint)i, Max = (uint)i, BestFit = true });
                    }

                    worksheetPart.Worksheet.AppendChild(columns);

                    worksheetPart.Worksheet.Save();

                    SheetData sheetData = new SheetData();
                    workbook.Workbook.Append(new Sheets());

                    Sheets sheets = workbook.Workbook.GetFirstChild<Sheets>();
                    if (sheets.Elements<Sheet>().Count() > 0)
                    {
                        sheets.RemoveAllChildren();
                    }

                    string relationshipId = workbook.GetIdOfPart(worksheetPart);
                    uint sheetId = 1;
                    string sheetName = "Sheet " + sheetId;

                    Sheet sheet = new Sheet()
                    {
                        Id = relationshipId,
                        SheetId = sheetId,
                        Name = sheetName
                    };

                    sheets.Append(sheet);

                    WorkbookStylesPart stylesPart = doc.WorkbookPart.AddNewPart<WorkbookStylesPart>();

                    Stylesheet stylesheet = new Stylesheet();

                    Fonts fonts = new Fonts();
                    fonts.AppendChild(new Font()
                    {
                        Bold = new Bold() { Val = false }
                    });
                    fonts.AppendChild(new Font()
                    {
                        Bold = new Bold() { Val = true }
                    });

                    Fills fills = new Fills();
                    fills.AppendChild(new Fill()
                    {
                        PatternFill = new PatternFill()
                        {
                            PatternType = PatternValues.None
                        }
                    });
                    fills.AppendChild(new Fill()
                    {
                        PatternFill = new PatternFill()
                        {
                            PatternType = PatternValues.Gray125
                        }
                    });
                    fills.AppendChild(new Fill()
                    {
                        PatternFill = new PatternFill()
                        {
                            PatternType = PatternValues.Solid,
                            ForegroundColor = new ForegroundColor() { Rgb = "FFF8F9FA" },
                            BackgroundColor = new BackgroundColor { Rgb = "FFF8F9FA" }
                        }
                    });

                    Borders borders = new Borders();
                    borders.AppendChild(new Border()
                    {
                        LeftBorder = new LeftBorder()
                        {
                            Style = BorderStyleValues.None
                        },
                        TopBorder = new TopBorder()
                        {
                            Style = BorderStyleValues.None
                        },
                        RightBorder = new RightBorder()
                        {
                            Style = BorderStyleValues.None
                        },
                        BottomBorder = new BottomBorder()
                        {
                            Style = BorderStyleValues.None
                        }
                    });
                    borders.AppendChild(new Border()
                    {
                        LeftBorder = new LeftBorder()
                        {
                            Color = new Color() { Rgb = "FFC1C1C1" },
                            Style = BorderStyleValues.Medium
                        },
                        TopBorder = new TopBorder()
                        {
                            Color = new Color() { Rgb = "FFC1C1C1" },
                            Style = BorderStyleValues.Medium
                        },
                        RightBorder = new RightBorder()
                        {
                            Color = new Color() { Rgb = "FFC1C1C1" },
                            Style = BorderStyleValues.Medium
                        },
                        BottomBorder = new BottomBorder()
                        {
                            Color = new Color() { Rgb = "FFC1C1C1" },
                            Style = BorderStyleValues.Medium
                        }
                    });

                    CellFormats cellformats = new CellFormats();
                    cellformats.AppendChild(new CellFormat()
                    {
                        FormatId = 0,
                        FillId = 0,
                        BorderId = 0,
                        Alignment = new Alignment()
                        {
                            Horizontal = HorizontalAlignmentValues.Left,
                            Vertical = VerticalAlignmentValues.Center
                        }
                    });
                    cellformats.AppendChild(new CellFormat()
                    {
                        FormatId = 1,
                        FillId = 2,
                        BorderId = 1,
                        Alignment = new Alignment()
                        {
                            Horizontal = HorizontalAlignmentValues.Center,
                            Vertical = VerticalAlignmentValues.Center
                        }
                    });

                    stylesheet.Append(fonts);
                    stylesheet.Append(fills);
                    stylesheet.Append(borders);
                    stylesheet.Append(cellformats);

                    stylesPart.Stylesheet = stylesheet;
                    stylesPart.Stylesheet.Save();

                    for (int i = 0; i < data.GetLength(0); i++)
                    {
                        Row newRow = new Row()
                        {
                            RowIndex = new DocumentFormat.OpenXml.UInt32Value((uint)i + 1)
                        };

                        for (int j = 0; j < data.GetLength(1); j++)
                        {
                            try
                            {
                                string columnName = Internal.GetExcelColumnName(j + 1);

                                object value = data[i, j];

                                if (value != null)
                                {
                                    Type type = data[i, j].GetType();
                                    Cell newCell;

                                    if (type == typeof(Int32) || type == typeof(Int16) || type == typeof(Int64) || type == typeof(Double) || type == typeof(Single) || type == typeof(Decimal))
                                    {
                                        newCell = new Cell() { CellValue = new CellValue(Convert.ToString(value)), DataType = CellValues.Number, CellReference = columnName.ToUpper() + newRow.RowIndex };
                                    }
                                    else if (type == typeof(DateTime))
                                    {
                                        newCell = new Cell() { CellValue = new CellValue((DateTime)value), DataType = CellValues.Date, CellReference = columnName.ToUpper() + newRow.RowIndex };
                                    }
                                    else if (type == typeof(Boolean))
                                    {
                                        newCell = new Cell() { CellValue = new CellValue((Boolean)value == false ? "0" : "1"), DataType = CellValues.Boolean, CellReference = columnName.ToUpper() + newRow.RowIndex };
                                    }
                                    else
                                    {
                                        newCell = new Cell() { CellValue = new CellValue(value.ToString()), DataType = CellValues.String, CellReference = columnName.ToUpper() + newRow.RowIndex };
                                    }

                                    if (i == 0 && headerFirstRow == true)
                                    {
                                        newCell.StyleIndex = 1;
                                    }
                                    else
                                    {
                                        newCell.StyleIndex = 0;
                                    }

                                    newRow.Append(newCell);
                                }
                            }
                            catch
                            {
                                continue;
                            }
                        }

                        if (newRow.Descendants<Cell>().Count() > 0)
                        {
                            sheetData.Append(newRow);
                        }
                    }

                    worksheetPart.Worksheet.AppendChild(sheetData);

                    worksheetPart.Worksheet.Save();

                    doc.Save();
                }

                stream.Flush();
                stream.Seek(0, SeekOrigin.Begin);
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }
    }
}
