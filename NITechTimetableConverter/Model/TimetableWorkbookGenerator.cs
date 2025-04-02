using ClosedXML.Excel;
using NITechTimetableConverter.Data;
using NITechTimetableConverter.Properties;

namespace NITechTimetableConverter.Model
{
    internal class TimetableWorkbookGenerator
    {
        private static readonly int[] dayOfWeekColumnNumbers = { 1, 12, 27, 38, 49 };
        private static readonly int dayOfWeekRowNumber = 5;
        public static XLWorkbook GenerateTimetableWorkbook(IEnumerable<IEnumerable<IEnumerable<Lecture>>> lectures)
        {
            XLWorkbook workbook = new();
            using (XLWorkbook template = new(new MemoryStream(Resources.XLSXTemplate)))
            {
                for (int i = 0; i < lectures.Count(); i++)
                {
                    IXLWorksheet worksheet = GenerateTimetableSheet(template, i.ToString(), lectures.ElementAt(i));
                    MergeContinuousCellsWithSameText(worksheet);
                    workbook.AddWorksheet(worksheet);
                }
                return workbook;
            }
        }

        private static void MergeContinuousCellsWithSameText(IXLWorksheet worksheet)
        {
            string lastCellValue = string.Empty;
            int startColumn = -1;
            int startRow = -1;
            int endColumn = -1;
            int endRow = -1;
            foreach (var row in worksheet.RowsUsed())
            {
                foreach (var cell in row.CellsUsed())
                {
                    string cellValue = cell.Value.ToString();

                    if (cell.IsMerged())
                    {
                        IXLRange mergedRange = cell.MergedRange();
                        IXLCell firstCell = mergedRange.FirstCell();
                        IXLCell lastCell = mergedRange.LastCell();
                        cellValue = firstCell.Value.ToString();
                        if (lastCellValue == cellValue && !string.IsNullOrEmpty(cellValue))
                        {
                            if (startColumn == -1)
                            {
                                startColumn = firstCell.WorksheetColumn().ColumnNumber();
                                startRow = firstCell.WorksheetRow().RowNumber();
                            }
                            endColumn = lastCell.WorksheetColumn().ColumnNumber();
                            endRow = lastCell.WorksheetRow().RowNumber();
                        }
                        else
                        {
                            if (startColumn != -1 && endColumn != -1 && startColumn != endColumn)
                            {
                                worksheet.Range(startRow, startColumn, endRow, endColumn).Merge();
                            }
                            startColumn = cell.WorksheetColumn().ColumnNumber();
                            startRow = cell.WorksheetRow().RowNumber();
                            endColumn = startColumn;
                            endRow = startRow;
                        }

                        lastCellValue = cellValue;
                    }

                }

                if (startColumn != -1 && endColumn != -1 && startColumn != endColumn)
                {
                    worksheet.Range(startRow, startColumn, endRow, endColumn).Merge();
                }

                lastCellValue = string.Empty;
                startColumn = -1;
                startRow = -1;
                endColumn = -1;
                endRow = -1;
            }
        }

        private static IXLWorksheet GenerateTimetableSheet(XLWorkbook template, string sheetName, IEnumerable<IEnumerable<Lecture>> lectures)
        {

            IXLWorksheet worksheet = template.Worksheet(1).CopyTo(sheetName);
            for (int i = 0; i < lectures.Count(); i++)
            {
                WriteCellByDayOfWeek(lectures, worksheet, i);
            }
            return worksheet;

        }

        private static void WriteCellByDayOfWeek(IEnumerable<IEnumerable<Lecture>> lectures, IXLWorksheet worksheet, int dayOfWeekIndex)
        {
            for (int ii = 0; ii < lectures.ElementAt(dayOfWeekIndex).Count(); ii++)
            {

                Lecture lecture = lectures.ElementAt(dayOfWeekIndex).OrderBy((l) => l.Classes.Length).ElementAt(ii);
                Array.Sort(lecture.Classes);
                int previousClassIndex = -1;
                int currentClassIndex = -1;
                for (int iii = 0; iii < lecture.Classes.Length; iii++)
                {
                    if (currentClassIndex == -1)
                    {
                        currentClassIndex = (int)lecture.Classes[iii];
                    }
                    int currentIndex = (int)lecture.Classes[iii];
                    if (previousClassIndex != currentIndex - 1 && previousClassIndex != -1)
                    {
                        WriteAndMergeCell(worksheet, dayOfWeekIndex, lecture, previousClassIndex, currentClassIndex);
                        currentClassIndex = currentIndex;
                    }
                    previousClassIndex = currentIndex;
                }
                if (currentClassIndex != -1)
                {
                    WriteAndMergeCell(worksheet, dayOfWeekIndex, lecture, previousClassIndex, currentClassIndex);
                }
            }
        }

        private static void WriteAndMergeCell(IXLWorksheet worksheet, int dayOfWeekIndex, Lecture lecture, int previousIndex, int currentStartCellIndex)
        {
            int startRow = dayOfWeekRowNumber + currentStartCellIndex;
            int startColumn = dayOfWeekColumnNumbers[dayOfWeekIndex] + (int)lecture.Period * 2;
            IXLCell cell = worksheet.Cell(startRow, startColumn);
            int lectureRowLength = previousIndex - currentStartCellIndex;
            if (string.IsNullOrEmpty(cell.Value.ToString()))
            {
                cell.Value = lecture.ToString();
                //worksheet.Range(startRow, startColumn, startRow + (previousIndex - currentStartCellIndex), startColumn + ((lecture.Period == Period.OnDemand) ? 0 : 1)).Merge();
                SearchUsedCellsAndMergeUnusedArea(worksheet, lecture.ToString(), startRow, startColumn, startRow + lectureRowLength, startColumn + ((lecture.Period == Period.OnDemand) ? 0 : 1));
            }
            else
            {
                Lecture? existLecture = null;
                XLCellValue value = cell.Value;
                if (Lecture.TryParse(value.ToString(), null, out existLecture))
                {
                    if (existLecture.Name == lecture.Name && existLecture.ID != lecture.ID)
                    {
                        cell.Value += string.Format("{0}{1}{0}{2}{3}", Environment.NewLine, lecture.ID, lecture.Instructor, (string.IsNullOrEmpty(lecture.Room) ? string.Empty : Environment.NewLine + lecture.Room));
                    }
                    else
                    {
                        if (string.IsNullOrEmpty(lecture.Room) && lecture.Period != Period.OnDemand)
                        {
                            cell.Value = cell.Value.ToString().Replace(Environment.NewLine, " ");
                            cell.Value += string.Format("{4}{0}{3}{1}{3}{2}", lecture.ID, lecture.Name, lecture.Instructor, " ", Environment.NewLine);
                        }
                        else
                        {
                            cell.Value += $"{Environment.NewLine}{Environment.NewLine}{lecture.ToString()}";
                        }
                    }
                }
                else
                {
                    if (string.IsNullOrEmpty(lecture.Room) && lecture.Period != Period.OnDemand)
                    {
                        cell.Value += string.Format("{4}{0}{3}{1}{3}{2}", lecture.ID, lecture.Name, lecture.Instructor, " ", Environment.NewLine);
                    }
                    else
                    {
                        cell.Value += $"{Environment.NewLine}{Environment.NewLine}{lecture.ToString()}";
                    }
                }
                if (lectureRowLength >= 1 && GetLowerCellOriginRow(worksheet, startColumn, startRow) - startRow == 1)
                {
                    WriteAndMergeCell(worksheet, dayOfWeekIndex, lecture, previousIndex, currentStartCellIndex + 1);
                }
            }
        }

        private static void SearchUsedCellsAndMergeUnusedArea(IXLWorksheet worksheet, string cellValue, int startRow, int startColumn, int lastRow, int lastColumn)
        {
            IEnumerable<IXLCell> usedCells = worksheet.Range(startRow, startColumn, lastRow, startColumn).CellsUsed().Where((e) => { return !(e.WorksheetRow().RowNumber() == startRow && e.WorksheetColumn().ColumnNumber() == startColumn); });
            string startCellValue = worksheet.Cell(startRow, startColumn).Value.ToString();
            if (usedCells.Count() == 0)
            {
                worksheet.Range(startRow, startColumn, lastRow, lastColumn).Merge();
                return;
            }
            int lastUsedRow = -1;
            for (int i = 0; i < usedCells.Count(); i++)
            {
                int cellRow = usedCells.ElementAt(i).WorksheetRow().RowNumber();
                if (lastUsedRow == -1)
                {
                    int up = GetUpperCellOriginRow(worksheet, lastColumn, cellRow);
                    worksheet.Range(startRow, startColumn, up, lastColumn).Merge();
                    lastUsedRow = cellRow;
                }
                else if (GetLowerCellOriginRow(worksheet, lastColumn, lastUsedRow) <= cellRow - 1)
                {
                    IXLRange mergeAreaRange = worksheet.Range(GetLowerCellOriginRow(worksheet, lastColumn, lastUsedRow), startColumn, cellRow - 1, lastColumn);
                    mergeAreaRange.Merge();
                    mergeAreaRange.FirstCell().WorksheetRow().Cell(startColumn).FormulaA1 = worksheet.Cell(startRow, startColumn).Address.ToString();
                    lastUsedRow = cellRow;
                }
                if (i == usedCells.Count() - 1)
                {
                    int row = GetLowerCellOriginRow(worksheet, lastColumn, cellRow);
                    if (row <= lastRow)
                    {
                        IXLRange mergeAreaRange = worksheet.Range(row, startColumn, lastRow, lastColumn);
                        mergeAreaRange.Merge();
                        mergeAreaRange.FirstCell().WorksheetRow().Cell(startColumn).FormulaA1 = worksheet.Cell(startRow, startColumn).Address.ToString();
                    }
                }
            }
        }

        private static int GetUpperCellOriginRow(IXLWorksheet worksheet, int lastColumn, int cellRow)
        {
            IXLCell cell = worksheet.Cell(cellRow - 1, lastColumn);
            return ((cell.IsMerged()) ? cell.MergedRange().FirstCell().WorksheetRow().RowNumber() : cellRow - 1);
        }

        private static int GetLowerCellOriginRow(IXLWorksheet worksheet, int lastColumn, int cellRow)
        {
            IXLCell cell = worksheet.Cell(cellRow, lastColumn);
            return ((cell.IsMerged()) ? cell.MergedRange().LastCell().WorksheetRow().RowNumber() + 1 : cellRow + 1);
        }
    }
}