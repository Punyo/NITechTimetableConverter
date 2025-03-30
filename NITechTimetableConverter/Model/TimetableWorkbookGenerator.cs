using ClosedXML.Excel;
using NITechTimetableConverter.Data;
using NITechTimetableConverter.Properties;

namespace NITechTimetableConverter.Model
{
    internal class TimetableWorkbookGenerator
    {
        private static readonly int[] dayOfWeekColumnNumbers = { 1, 12, 27, 38, 49 };
        private static readonly int dayOfWeekRowNumber = 5;
        public static XLWorkbook GenerateTimetableWorkbook(IEnumerable<IEnumerable<Lecture>> lectures)
        {

            XLWorkbook returnWorkbook = new();
            using (XLWorkbook templateWorkbook = new(new MemoryStream(Resources.XLSXTemplate)))
            {
                IXLWorksheet worksheet = templateWorkbook.Worksheet(1);
                for (int i = 0; i < lectures.Count(); i++)
                {
                    WriteCellByDayOfWeek(lectures, worksheet, i);
                }
                returnWorkbook.AddWorksheet(worksheet);
            }
            return returnWorkbook;
        }

        private static void WriteCellByDayOfWeek(IEnumerable<IEnumerable<Lecture>> lectures, IXLWorksheet worksheet, int dayOfWeekIndex)
        {
            for (int ii = 0; ii < lectures.ElementAt(dayOfWeekIndex).Count(); ii++)
            {
                Lecture lecture = lectures.ElementAt(dayOfWeekIndex).ElementAt(ii);
                Array.Sort(lecture.Classes);
                int previousIndex = -1;
                int currentStartCellIndex = -1;
                for (int iii = 0; iii < lecture.Classes.Length; iii++)
                {
                    if (currentStartCellIndex == -1)
                    {
                        currentStartCellIndex = (int)lecture.Classes[iii];
                    }
                    int currentIndex = (int)lecture.Classes[iii];
                    if (previousIndex != currentIndex - 1 && previousIndex != -1)
                    {
                        WriteAndMergeCell(worksheet, dayOfWeekIndex, lecture, previousIndex, currentStartCellIndex);
                        currentStartCellIndex = currentIndex;
                    }
                    previousIndex = currentIndex;
                }
                if (currentStartCellIndex != -1)
                {
                    WriteAndMergeCell(worksheet, dayOfWeekIndex, lecture, previousIndex, currentStartCellIndex);
                }
            }
        }

        private static void WriteAndMergeCell(IXLWorksheet worksheet, int dayOfWeekIndex, Lecture lecture, int previousIndex, int currentStartCellIndex)
        {
            int startRow = dayOfWeekRowNumber + currentStartCellIndex;
            int startColumn = dayOfWeekColumnNumbers[dayOfWeekIndex] + (int)lecture.Period * 2;
            IXLCell cell = worksheet.Cell(startRow, startColumn);
            cell.Value = lecture.ToString();
            worksheet.Range(startRow, startColumn, startRow + (previousIndex - currentStartCellIndex), startColumn + (cell.IsMerged() ? 1 : 0)).Merge();
        }
    }
}