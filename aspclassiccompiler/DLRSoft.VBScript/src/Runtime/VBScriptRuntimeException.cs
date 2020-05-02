using System;
using System.Collections.Generic;
using System.Text;

namespace Dlrsoft.VBScript.Runtime
{
    public class VBScriptRuntimeException : Exception
    {
        private int _number;
        private string _source = string.Empty;
        private string _helpContext = string.Empty;
        private string _helpFile = string.Empty;

        public VBScriptRuntimeException()
        {
        }

        public VBScriptRuntimeException(Exception ex)
            : base(ex.Message, ex)
        {
            this._number = 507; //An exception occurred
        }

        public VBScriptRuntimeException(Exception ex, string source)
            : base(source, ex)
        {
            this._source = source;
        }

        public VBScriptRuntimeException(int number, string description)
            : this(number, description, null, null, null)
        {
        }

        public VBScriptRuntimeException(int number, string description, string source, string helpFile, string helpContext)
            : base(description)
        {
            this._number = number;
            this._source = source;
            this._helpContext = helpContext;
            this._helpFile = helpFile;
        }

        public int Number { get { return _number; } }
        public string Description { get { return base.Message; } }
        public string Source { get { return _source; } }
        public string HelpContext { get { return _helpContext; } }
        public string HelpFile { get { return _helpFile; } }
    }
}
