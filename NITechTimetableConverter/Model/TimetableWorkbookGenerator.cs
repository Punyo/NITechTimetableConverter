using ClosedXML.Excel;
using NITechTimetableConverter.Data;
using NITechTimetableConverter.Properties;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

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
                    foreach (Lecture lecture in lectures.ElementAt(i))
                    {
                        foreach (var c in lecture.Classes)
                        {
                            worksheet.Cell(dayOfWeekRowNumber + (int)c, dayOfWeekColumnNumbers[i] + (int)lecture.Period * 2).Value = lecture.ToString();
                        }
                    }
                }
                returnWorkbook.AddWorksheet(worksheet);
            }
            return returnWorkbook;
        }
    }
}