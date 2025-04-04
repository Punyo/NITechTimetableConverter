﻿using NITechTimetableConverter.Data;
using ClosedXML.Excel;
using NITechTimetableConverter.Properties;
using System.Diagnostics;
using NITechTimetableConverter.Utility;

namespace NITechTimetableConverter.Model
{
    internal class GenerateTimetableFromDegradedFormatModel(string degradedXLSXFilePath)
    {
        public void GenerateTimetable(string outputPath, bool openGeneratedFile = true)
        {
            Console.WriteLine(Resources.MessageDivider + Environment.NewLine);
            Console.WriteLine(Resources.MessageStartConvert);
            Console.WriteLine(Resources.MessageExtractingLectureInfo);
            IEnumerable<IEnumerable<IEnumerable<Lecture>>> lectureInfo = LectureExtractor.ExtractLecturesFromXLSXFile(degradedXLSXFilePath);
            Console.WriteLine(Resources.MessageExtractingLectureInfoComplete);
            Console.WriteLine(Resources.MessageWritingLectureInfo);
            XLWorkbook workbook = TimetableWorkbookGenerator.GenerateTimetableWorkbook(lectureInfo);
            workbook.Properties.Author = Resources.MessageCredit;
            workbook.SaveAs(outputPath);
            Console.WriteLine(Resources.MessageWritingLectureInfoComplete);
            ConsoleUtil.WriteLineWithColor(string.Format(Resources.MessageConvertComplete), ConsoleUtil.ColorMode.Info);
            if (openGeneratedFile)
            {
                Console.WriteLine(Resources.MessageOpeningConvertedXLSXFile);
                OpenXLSXFile(outputPath);
            }
        }

        private void OpenXLSXFile(string path)
        {
            ProcessStartInfo result = new()
            {
                FileName = path,
                UseShellExecute = true
            };
            Process.Start(result);
        }
    }
}