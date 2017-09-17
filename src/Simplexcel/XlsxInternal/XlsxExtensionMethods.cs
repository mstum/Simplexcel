using System;
namespace Simplexcel.XlsxInternal
{
    internal static class XlsxEnumFormatter
    {
        internal static string GetXmlValue(this Panes pane)
        {
            switch (pane)
            {
                case Panes.BottomLeft:
                    return "bottomLeft";
                case Panes.BottomRight:
                    return "bottomRight";
                case Panes.TopLeft:
                    return "topLeft";
                case Panes.TopRight:
                    return "topRight";
                default:
                    throw new ArgumentOutOfRangeException(nameof(pane), "Invalid Pane: " + pane);
            }
        }

        internal static string GetXmlValue(this PaneState state)
        {
            switch (state)
            {
                case PaneState.Split:
                    return "split";
                case PaneState.Frozen:
                    return "frozen";
                case PaneState.FrozenSplit:
                    return "frozenSplit";
                default:
                    throw new ArgumentOutOfRangeException(nameof(state), "Invalid Pane State: " + state);
            }
        }

        internal static double ToOleAutDate(this DateTime dt)
        {
#if NET45
            return dt.ToOADate();
#else
            long value = dt.Ticks;

            const long TicksPerMillisecond = 10000;
            const long TicksPerSecond = TicksPerMillisecond * 1000;
            const long TicksPerMinute = TicksPerSecond * 60;
            const long TicksPerHour = TicksPerMinute * 60;
            const long TicksPerDay = TicksPerHour * 24;

            const int MillisPerSecond = 1000;
            const int MillisPerMinute = MillisPerSecond * 60;
            const int MillisPerHour = MillisPerMinute * 60;
            const int MillisPerDay = MillisPerHour * 24;

            const int DaysPerYear = 365;
            const int DaysPer4Years = DaysPerYear * 4 + 1;
            const int DaysPer100Years = DaysPer4Years * 25 - 1;
            const int DaysPer400Years = DaysPer100Years * 4 + 1;
            const int DaysTo1899 = DaysPer400Years * 4 + DaysPer100Years * 3 - 367;

            const long DoubleDateOffset = DaysTo1899 * TicksPerDay;
            const long OADateMinAsTicks = (DaysPer100Years - DaysPerYear) * TicksPerDay;

            if (value == 0)
            {
                return 0.0;
            }
            if (value < TicksPerDay)
            {
                value += DoubleDateOffset;
            }
            if (value < OADateMinAsTicks)
            {
                throw new OverflowException("Not a legal OleAut date.");
            }

            long millis = (value - DoubleDateOffset) / TicksPerMillisecond;
            if (millis < 0)
            {
                long frac = millis % MillisPerDay;
                if (frac != 0) millis -= (MillisPerDay + frac) * 2;
            }
            return (double)millis / MillisPerDay;
#endif
        }
    }
}