using ClosedXML.Excel;
using NITechTimetableConverter.Data;
using System.Text.RegularExpressions;

namespace NITechTimetableConverter.Model
{
    internal class LectureExtractor
    {
        private readonly static string[] departmentAbbreviates = { "LC", "PE", "EM", "CS", "AC", "CR" };
        private readonly static string allDepartments = "全学科";
        private readonly static Lecture noInfoInWorksheet = new() { ID = "XXXX", Name = "シラバスを参照してください", Instructor = "XXXX", Classes = [.. Enum.GetValues<Classes>().Cast<Classes>()], Period = Period.OneTwo, Room = "XXXX" };


        public static IEnumerable<IEnumerable<Lecture>> ExtractLecturesFromXLSXFile(string path, int worksheetIndex = 1)
        {
            using (var workbook = new XLWorkbook(path))
            {
                var worksheet = workbook.Worksheet(worksheetIndex);
                int dayOfWeekColumnNumber = worksheet.Cell("RowName").WorksheetColumn().ColumnNumber();
                IEnumerable<IXLRange> dayOfWeekRange = worksheet.MergedRanges.Where(r => r.RangeAddress.FirstAddress.ColumnNumber == dayOfWeekColumnNumber);
                return dayOfWeekRange.Select(r => GetLecturesByDayOfWeekRangeInRowName(worksheet, r)).ToArray();
            }
        }

        private static IEnumerable<Lecture> GetLecturesByDayOfWeekRangeInRowName(IXLWorksheet worksheet, IXLRange dayOfWeekRangeInRowName)
        {
            IXLRangeAddress rangeAddress = dayOfWeekRangeInRowName.RangeAddress;
            List<Lecture> lectures = new();
            for (int j = 0; j <= 5; j++)
            {
                for (int i = rangeAddress.FirstAddress.RowNumber; i <= rangeAddress.LastAddress.RowNumber; i++)
                {
                    int baseColumnNumber = j * 6 + rangeAddress.FirstAddress.ColumnNumber;
                    string id = worksheet.Cell(i, baseColumnNumber + 1).Value.ToString();
                    string commaSeperatedClasses = worksheet.Cell(i, baseColumnNumber + (j == 5 ? 4 : 5)).Value.ToString();
                    string lectureName = worksheet.Cell(i, baseColumnNumber + 2).Value.ToString();
                    if (commaSeperatedClasses.IndexOf("技術と歴史・哲学") > 0 || commaSeperatedClasses.IndexOf("技術と社会・国際") > 0 ||
                        commaSeperatedClasses.IndexOf("技術と人間・心理") > 0 || (lectureName.IndexOf("English") > -1 && lectureName.IndexOf("Global") == -1))
                    {
                        lectures.Add(noInfoInWorksheet with { Period = (Period)j });
                        break;
                    }
                    if (string.IsNullOrEmpty(id))
                    {
                        break;
                    }
                    if (string.IsNullOrEmpty(commaSeperatedClasses) || commaSeperatedClasses.IndexOf("技術と歴史・哲学") == 0 ||
                        commaSeperatedClasses.IndexOf("技術と社会・国際") == 0 || commaSeperatedClasses.IndexOf("技術と人間・心理") == 0)
                    {
                        commaSeperatedClasses = allDepartments;
                    }
                    Lecture lecture = new Lecture
                    {
                        ID = id,
                        Name = lectureName,
                        Instructor = worksheet.Cell(i, baseColumnNumber + 3).Value.ToString(),
                        Room = j == 5 ? null : worksheet.Cell(i, baseColumnNumber + 4).Value.ToString(),
                        Period = (Period)j,
                        Classes = ConvertCommaSeperatedStringToClasses(commaSeperatedClasses)
                    };
                    lectures.Add(lecture);
                }
            }
            return lectures;
        }

        private static Classes[] ConvertCommaSeperatedStringToClasses(string s)
        {
            string[] classStrings = s.Split(',').Select(s => s.Replace("◆", "").Replace("前半", "First").Replace("後半", "Second").Trim()).ToArray();
            List<Classes> classes = new List<Classes>();
            for (int i = 0; i < classStrings.Length; i++)
            {
                if (string.IsNullOrEmpty(classStrings[i]))
                {
                    continue;
                }
                try
                {
                    classes.Add((Classes)Enum.Parse(typeof(Classes), classStrings[i]));
                }
                catch (ArgumentException)
                {
                    if (classStrings[i] == allDepartments)
                    {
                        classes.AddRange(Enum.GetValues<Classes>().Cast<Classes>());
                    }
                    else if (departmentAbbreviates.Any((s) => { return classStrings[i].Contains(s); }))
                    {
                        classStrings[i] = Regex.Replace(classStrings[i], "[^a-zA-Z]", "");
                        classes.AddRange(Enum.GetNames<Classes>().Where(s => s.Contains(classStrings[i])).Select(Enum.Parse<Classes>));
                    }
                }
            }
            return classes.ToArray();
        }
    }
}
