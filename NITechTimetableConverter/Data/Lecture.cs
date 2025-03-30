using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace NITechTimetableConverter.Data
{
    public record Lecture
    {
        public required string ID { get; init; }
        public required string Name { get; init; }
        public required string Instructor { get; init; }
        public required Classes[] Classes { get; init; }
        public required Period Period { get; init; }
        public string? Room { get; init; }

        public override string ToString()
        {
            return $"{ID}{Environment.NewLine}{Name}{Environment.NewLine}{Instructor}{Environment.NewLine}{Room}";
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
