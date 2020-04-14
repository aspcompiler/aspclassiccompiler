using System;
using System.Collections;
using System.Collections.Generic;
using System.Text;
using Microsoft.Scripting;

namespace Dlrsoft.VBScript.Compiler
{
    public class SourceUtil
    {
        public static int GetLineCount(string source)
        {
            int lineCount = 0;
            if (string.IsNullOrEmpty(source))
                return lineCount;

            int start = 0;
            int len = source.Length;
            while (start < len)
            {
                int pos = source.IndexOf('\r', start);
                if (pos > -1)
                {
                    lineCount++;
                }
                else
                {
                    break;
                }

                start = pos + 1;
                if (start < len && source[start] == '\n')
                {
                    start++;
                }
            }

            //Last line with CR/LF
            //We got an empty line in this situation
            if (len >= start)
            {
                lineCount++;
            }

            return lineCount;
        }

        public static Range[] GetLineRanges(string source)
        {
            if (string.IsNullOrEmpty(source))
                return new Range[]{};

            List<Range> lineRanges = new List<Range>();
            int start = 0;
            int len = source.Length;

            while (start < len)
            {
                int pos = source.IndexOf('\r', start);
                if (pos < 0)
                {
                    break;
                }

                if (pos + 1 < len && source[pos + 1] == '\n')
                {
                    pos++;
                }

                lineRanges.Add(new Range(start, pos));
                start = pos + 1;
            }

            //Last line with CR/LF
            if (len > start)
            {
                lineRanges.Add(new Range(start, len-1));
            }

            return lineRanges.ToArray();
        }

        public static LineColumn GetLineColumn(Range[] lineRanges, int index)
        {
            IComparer comparer = new RangeComparer();
            int i = Array.BinarySearch(lineRanges, 0, lineRanges.Length, index, comparer);
            return new LineColumn(i + 1, index - lineRanges[i].Start + 1);
        }

        public static SourceSpan GetSpan(Range[] lineRanges, int start, int end)
        {
            LineColumn lcStart = GetLineColumn(lineRanges, start);
            SourceLocation slStart = new SourceLocation(start, lcStart.Line, lcStart.Column);
            LineColumn lcEnd = GetLineColumn(lineRanges, end);
            SourceLocation slEnd = new SourceLocation(end, lcEnd.Line, lcEnd.Column);
            return new SourceSpan(slStart, slEnd);
        }

        internal static SourceSpan ConvertSpan(Dlrsoft.VBScript.Parser.Span span)
        {
            Dlrsoft.VBScript.Parser.Location start = span.Start;
            Dlrsoft.VBScript.Parser.Location end = span.Finish;
            return new SourceSpan(
                new SourceLocation(start.Index, start.Line, start.Column),
                new SourceLocation(end.Index, end.Line, end.Column)
            );
        }
    }

    public struct Range
    {
        public int Start;
        public int End;

        public Range(int start, int end)
        {
            this.Start = start;
            this.End = end;
        }
    }

    public struct LineColumn
    {
        public int Line;
        public int Column;
        
        public LineColumn(int line, int column)
        {
            this.Line = line;
            this.Column = column;
        }
    }

    internal class RangeComparer : IComparer
    {
        #region IComparer Members

        public int Compare(object range, object index)
        {
            Range theRange = (Range)range;
            int theIndex = (int)index;

            if (theRange.End < theIndex)
            {
                return -1;
            }
            else if (theRange.Start > theIndex)
            {
                return 1;
            }
            return 0;
        }

        #endregion
    }

    internal class SpanComparer : IComparer<SourceSpan>
    {
        #region IComparer Members

        public int Compare(SourceSpan x, SourceSpan y)
        {
            if (x.End.Line < y.Start.Line)
            {
                return -1;
            }
            else if (x.Start.Line > y.End.Line)
            {
                return 1;
            }
            else
            {
                return 0;
            }
        }

        #endregion
    }
}
