using NITechTimetableConverter.Data;
using ClosedXML.Excel;

namespace NITechTimetableConverter.Model
{
    internal class GenerateTimetableFromDegradedFormatModel(string degradedXLSXFilePath)
    {
        public void GenerateTimetable(string outputPath, int worksheetIndex = 1)
        {
            XLWorkbook workbook = TimetableWorkbookGenerator.GenerateTimetableWorkbook(LectureExtractor.ExtractLecturesFromXLSXFile(degradedXLSXFilePath, worksheetIndex));
            workbook.SaveAs(outputPath);
        }
    }
}