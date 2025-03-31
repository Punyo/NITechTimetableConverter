namespace NITechTimetableConverter.Utility
{
    internal class ConsoleUtil
    {
        public static void WriteWithColor(string message, ColorMode color)
        {
            switch (color)
            {
                case ColorMode.Info:
                    Console.ForegroundColor = ConsoleColor.Green;
                    break;
                case ColorMode.Warning:
                    Console.ForegroundColor = ConsoleColor.Yellow;
                    break;
                case ColorMode.Error:
                    Console.ForegroundColor = ConsoleColor.Red;
                    break;
            }
            Console.Write(message);
            Console.ResetColor();
        }

        public static void WriteLineWithColor(string message, ColorMode color)
        {
            WriteWithColor(message + Environment.NewLine, color);
        }

        public enum ColorMode
        {
            Info,
            Warning,
            Error
        }
    }
}
