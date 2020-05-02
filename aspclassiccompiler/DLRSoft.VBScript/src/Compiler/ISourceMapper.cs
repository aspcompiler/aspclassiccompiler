using System;
using System.Collections.Generic;
using System.Text;
using Microsoft.Scripting;

namespace Dlrsoft.VBScript.Compiler
{
    public interface ISourceMapper
    {
        DocSpan Map(SourceSpan span);
    }
}
