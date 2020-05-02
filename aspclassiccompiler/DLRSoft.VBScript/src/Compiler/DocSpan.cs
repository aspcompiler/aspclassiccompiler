using System;
using System.Collections.Generic;
using System.Text;
using Microsoft.Scripting;

namespace Dlrsoft.VBScript.Compiler
{
    public class DocSpan
    {
        public DocSpan(string uri, SourceSpan span)
        {
            this.Uri = uri;
            this.Span = span;
        }

        public string Uri { get; set; }
        public SourceSpan Span { get; set; }
    }
}
