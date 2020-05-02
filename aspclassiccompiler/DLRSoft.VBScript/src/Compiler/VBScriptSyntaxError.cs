using System;
using System.Collections.Generic;
using System.Text;
using Microsoft.Scripting;

namespace Dlrsoft.VBScript.Compiler
{
    public class VBScriptSyntaxError
    {
        public VBScriptSyntaxError(string fileName, SourceSpan span, int errCode, string errorDesc)
        {
            this.FileName = fileName;
            this.Span = span;
            this.ErrorCode = errCode;
            this.ErrorDescription = errorDesc;
        }

        public string FileName { get; set; }
        public SourceSpan Span { get; set; }
        public int ErrorCode { get; set; }
        public string ErrorDescription { get; set; }
    }
}
