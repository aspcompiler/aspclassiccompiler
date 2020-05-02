using System;
using System.Collections.Generic;
using System.Text;

namespace Dlrsoft.VBScript.Compiler
{
    public class VBScriptCompilerException : Exception
    {
        private IList<VBScriptSyntaxError> _errors;

        public VBScriptCompilerException(IList<VBScriptSyntaxError> errors)
        {
            _errors = errors;
        }

        public IList<VBScriptSyntaxError> SyntaxErrors
        {
            get { return _errors; }
        }
    }
}
