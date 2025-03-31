using System.Diagnostics.CodeAnalysis;

namespace NITechTimetableConverter.Data
{
    public record Lecture : IParsable<Lecture>
    {
        public required string ID { get; init; }
        public required string Name { get; init; }
        public required string Instructor { get; init; }
        public required Classes[] Classes { get; init; }
        public required Period Period { get; init; }
        public string? Room { get; init; }

        public static Lecture Parse(string s, IFormatProvider? provider = null)
        {
            string[] strings = s.Split(Environment.NewLine);
            return new Lecture
            {
                ID = strings[0],
                Name = strings[1],
                Instructor = strings[2],
                Room = (strings.Length > 3 ? strings[3] : null),
                Classes = Array.Empty<Classes>(),
                Period = Period.OneTwo
            };
        }

        public static bool TryParse([NotNullWhen(true)] string? s, IFormatProvider? provider, [MaybeNullWhen(false)] out Lecture result)
        {
            try
            {
                result = Parse(s);
                return true;
            }
            catch (Exception e) when (e is NullReferenceException || e is IndexOutOfRangeException)
            {
                result = null;
                return false;
            }
        }

        public override string ToString()
        {
            if (string.IsNullOrEmpty(Room))
            {
                return $"{ID}{Environment.NewLine}{Name}{Environment.NewLine}{Instructor}";
            }
            else
            {
                return $"{ID}{Environment.NewLine}{Name}{Environment.NewLine}{Instructor}{Environment.NewLine}{Room}";
            }
        }

    }

    public enum Classes
    {
        LCa,
        LCb,
        LCc,
        PEa,
        PEb,
        EMaFirst,
        EMaSecond,
        EMbFirst,
        EMbSecond,
        CSaFirst,
        CSaSecond,
        CSbFirst,
        CSbSecond,
        CScFirst,
        CScSecond,
        ACa,
        ACc,
        ACm,
        CR1,
        CR2,
        CR3,
        CR4,
        CR5,
        CR6,
        CR7,
        CR8,
        CR9,
        CR10,
        CR11,
        CR12,
        CR13
    }
}
