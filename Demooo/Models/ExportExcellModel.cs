using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Web;
using System.Web.Mvc;
using System.Xml.Linq;

namespace Demooo.Models {
    
    public class ExcelCellMeta {
        public string Content { get; set; }
        public string StyleName { get; set; }
        public int ColumnIndex { get; set; }
        public int ColSpan { get; set; }
        public int RowSpan { get; set; }
    }

    public class ExcelMeta {
        public List<List<ExcelCellMeta>> Meta { get; set; }
        public double[] ColumnWidths { get; set; }
    }
    
    public class ExcelModel {
        [AllowHtml]
        public string Data { get; set; }
    }

    public class ExcelHelper {
        IWorkbook workbook;
        public ICellStyle heading1 { get; set; }
        public ICellStyle heading2 { get; set; }
        public ICellStyle rowHead { get; set; }
        public ICellStyle columnHead { get; set; }
        public ICellStyle content { get; set; }
        IFont fontWhite;
        IFont fontBlack;
        ISheet sheet;

        public ExcelMeta GetExcelMeta(string theTableHtml) {
            ExcelMeta excelMeta = new ExcelMeta();
            List<List<ExcelCellMeta>> meta = new List<List<ExcelCellMeta>>();
            double[] columnWidths;
            var theTable = XElement.Parse(theTableHtml);
            if (theTable.Name.LocalName.ToLower() == "table") {
                int numberOfColumns = 0;
                if (theTable.Attributes().Any(a => a.Name.LocalName == "data-xls-columns")) {
                    if (!int.TryParse(theTable.Attribute("data-xls-columns").Value, out numberOfColumns)) {
                        throw new Exception("'data-xls-Columns' on table element must have a numeric value.");
                    }

                    columnWidths = new double[numberOfColumns];
                    if (theTable.Attributes().Any(a => a.Name.LocalName == "data-xls-column-widths")) {
                        string[] temp = theTable.Attribute("data-xls-column-widths").Value.Split(',');
                        double w;
                        for (int i = 0; i < temp.Length; i++) {
                            if (double.TryParse(temp[i], out w)) {
                                columnWidths[i] = w;
                            }

                        }
                        //columnWidths.Where(c => c == 0).ToList().ForEach(c => c = 1);
                    }
                    excelMeta.ColumnWidths = columnWidths;
                    List<XElement> rowList = null;
                    if (theTable.Elements().Any(x => x.Name.LocalName.ToLower() == "tbody"))
                        rowList = theTable.Elements().Where(t => t.Name.LocalName.ToLower() == "tbody").FirstOrDefault().Elements().Where(
                        x => x.Name.LocalName.ToLower() == "tr" && (!x.Attributes().Any(a => a.Name.LocalName == "data-xls-exclude") || x.Attribute("data-xls-exclude").Value.ToLower() != "true")
                        ).ToList();
                    else
                        rowList = theTable.Elements().Where(
                        x => x.Name.LocalName.ToLower() == "tr" && (!x.Attributes().Any(a => a.Name.LocalName == "data-xls-exclude") || x.Attribute("data-xls-exclude").Value.ToLower() != "true")
                        ).ToList();
                    if (rowList.Count() == 0)
                        throw new Exception("No rows found.");
                    rowList.ForEach(r => {
                        List<ExcelCellMeta> rowMeta = new List<ExcelCellMeta>();
                        var columnList = r.Elements().Where(x => x.Name.LocalName.ToLower() == "td").ToList();
                        columnList.ForEach(c => {
                            //int width = c.Attributes().Any(a => a.Name.LocalName == "data-xls-width") ? int.Parse(c.Attribute("data-xls-width").Value) : 1;

                            rowMeta.Add(new ExcelCellMeta()
                            {
                                Content = c.Value,
                                StyleName = c.Attributes().Any(a => a.Name.LocalName == "data-xls-class") ? c.Attribute("data-xls-class").Value : "",
                                //Width = c.Attributes().Any(a => a.Name.LocalName == "data-xls-width") ? int.Parse(c.Attribute("data-xls-width").Value) : 1,
                                ColSpan = c.Attributes().Any(a => a.Name.LocalName.ToLower() == "colspan") ? int.Parse(c.Attributes().Where(a => a.Name.LocalName.ToLower() == "colspan").FirstOrDefault().Value) : 1,
                                RowSpan = c.Attributes().Any(a => a.Name.LocalName.ToLower() == "rowspan") ? int.Parse(c.Attributes().Where(a => a.Name.LocalName.ToLower() == "rowspan").FirstOrDefault().Value) : 1,
                                ColumnIndex = c.Attributes().Any(a => a.Name.LocalName.ToLower() == "data-xls-col-index") ? int.Parse(c.Attributes().Where(a => a.Name.LocalName.ToLower() == "data-xls-col-index").FirstOrDefault().Value) : -1
                            });
                        });
                        meta.Add(rowMeta);
                    });

                }
                else {
                    throw new Exception("Please provide 'data-xls-columns' attribute on the table element to qualify it with numner of columns.");
                }

            }
            else {
                throw new Exception("Provided Html is not that of a table element.");
            }
            excelMeta.Meta = meta;
            return excelMeta;
        }
        public byte[] GetExcelDocument(ExcelMeta excelMeta) {

            workbook = new XSSFWorkbook();


            sheet = workbook.CreateSheet("CARS");
            fontWhite = workbook.CreateFont(); fontWhite.Color = IndexedColors.White.Index; fontWhite.IsBold = true;
            fontBlack = workbook.CreateFont(); fontBlack.IsBold = true;

            heading1 = GetBasicHeaderStyle(IndexedColors.Grey80Percent.Index, fontWhite, IndexedColors.Grey80Percent.Index);
            heading2 = GetBasicHeaderStyle(IndexedColors.Grey40Percent.Index, fontBlack, IndexedColors.Grey80Percent.Index);
            heading2.Alignment = HorizontalAlignment.Left;
            rowHead = GetBasicHeaderStyle(IndexedColors.White.Index, fontBlack, IndexedColors.Grey80Percent.Index);
            columnHead = rowHead;

            content = GetBasicStyle();

            for (int index = 0; index < excelMeta.ColumnWidths.Length; index++) {
                sheet.SetColumnWidth(index, Convert.ToInt32(256 * 20 * (excelMeta.ColumnWidths[index] > 0 ? excelMeta.ColumnWidths[index] : 1)));
            }
            for (int r = 0; r < excelMeta.Meta.Count; r++) {
                IRow row = sheet.GetRow(r) == null ? sheet.CreateRow(r) : sheet.GetRow(r);
                List<ExcelCellMeta> cList = excelMeta.Meta.ElementAt(r);
                for (int c = 0; c < cList.Count; c++) {
                    List<ICell> placeHoldres = row.Cells;
                    ExcelCellMeta cellMeta = cList.ElementAt(c);
                    if (!placeHoldres.Any(pc => pc.ColumnIndex == cellMeta.ColumnIndex)) {

                        ICell cell = row.CreateCell(cellMeta.ColumnIndex);
                        if (!string.IsNullOrEmpty(cellMeta.Content))
                            cell.SetCellValue(cellMeta.Content);
                        if (GetPropValue(cellMeta.StyleName) != null && GetPropValue(cellMeta.StyleName) is ICellStyle) {
                            cell.CellStyle = GetPropValue(cellMeta.StyleName) as ICellStyle;
                        }
                        else {
                            cell.CellStyle = content;
                        }
                        int spanUptoColumn = cellMeta.ColumnIndex + cellMeta.ColSpan - 1;
                        int spanUptoRow = r + cellMeta.RowSpan - 1;
                        if (cellMeta.ColSpan > 1 && cellMeta.RowSpan <= 1) {
                            for (int s = 1; s < cellMeta.ColSpan; s++) {
                                ICell dummy = row.CreateCell(cellMeta.ColumnIndex + s);
                                dummy.CellStyle = cell.CellStyle;

                            }
                        }
                        else if (cellMeta.RowSpan > 1 && cellMeta.ColSpan <= 1) {

                            for (int s = 1; s < cellMeta.RowSpan; s++) {
                                IRow dummyRow = sheet.CreateRow(r + s);
                                ICell dummy = dummyRow.CreateCell(cellMeta.ColumnIndex);
                                dummy.CellStyle = cell.CellStyle;

                            }
                        }
                        else if (cellMeta.RowSpan > 1 && cellMeta.ColSpan > 1) {
                            for (int cs = 0; cs < cellMeta.ColSpan; cs++) {
                                if (cs != 0) {
                                    ICell dummy = row.CreateCell(cellMeta.ColumnIndex + cs);
                                    dummy.CellStyle = cell.CellStyle;
                                }
                                for (int rs = 1; rs < cellMeta.RowSpan; rs++) {
                                    IRow dummyRow = sheet.GetRow(r + rs) == null ? sheet.CreateRow(r + rs) : sheet.GetRow(r + rs);
                                    ICell dummy = dummyRow.CreateCell(cellMeta.ColumnIndex + cs);
                                    dummy.CellStyle = cell.CellStyle;
                                }
                            }
                        }
                        if (cellMeta.RowSpan > 1 || cellMeta.ColSpan > 1) {
                            NPOI.SS.Util.CellRangeAddress cra = new NPOI.SS.Util.CellRangeAddress(r, spanUptoRow, cellMeta.ColumnIndex, spanUptoColumn);
                            sheet.AddMergedRegion(cra);
                        }

                    }

                }
            }
            MemoryStream ms = new MemoryStream();
            workbook.Write(ms);

            byte[] toReturn = ms.ToArray();
            return toReturn;

        }

        private ICellStyle GetBasicHeaderStyle(short backColor, IFont font, short borderColor) {
            ICellStyle basicStyle = workbook.CreateCellStyle();
            basicStyle.FillPattern = FillPattern.SolidForeground;
            basicStyle.FillForegroundColor = backColor;
            basicStyle.SetFont(font);
            basicStyle.BorderRight = NPOI.SS.UserModel.BorderStyle.Thin;
            basicStyle.BorderBottom = NPOI.SS.UserModel.BorderStyle.Thin;
            basicStyle.RightBorderColor = borderColor;
            basicStyle.BottomBorderColor = borderColor;
            basicStyle.Alignment = HorizontalAlignment.Center;
            return basicStyle;
        }
        private ICellStyle GetBasicStyle() {
            ICellStyle basicStyle = workbook.CreateCellStyle();
            basicStyle.WrapText = true;
            basicStyle.BorderRight = NPOI.SS.UserModel.BorderStyle.Thin;
            basicStyle.BorderBottom = NPOI.SS.UserModel.BorderStyle.Thin;
            basicStyle.RightBorderColor = IndexedColors.Grey80Percent.Index;
            basicStyle.BottomBorderColor = IndexedColors.Grey80Percent.Index;

            return basicStyle;
        }
        private object GetPropValue(string propName) {
            try {
                return this.GetType().GetProperty(propName).GetValue(this, null);
            }
            catch {
                return null;
            }
        }
    }

    public class ExportViewModel {
        [AllowHtml]
        public string Csv { get; set; }
    }

}