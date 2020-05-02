using System;
using System.Collections.Generic;
using System.Text;

namespace Dlrsoft.VBScript.Runtime
{
    public class ErrObject
    {
        private int _number;
        private string _description = string.Empty;
        private string _source = string.Empty;
        private string _helpContext = string.Empty;
        private string _helpFile = string.Empty;

        public int Number { get { return _number; } }
        public string Description { get { return _description; } }
        public string Source { get { return _source; } }
        public string HelpContext { get { return _helpContext; } }
        public string HelpFile { get { return _helpFile; } }

        public void Clear()
        {
            if (_number != 0)
            {
                _number = 0;
                _description = string.Empty;
                _source = string.Empty;
                _helpContext = string.Empty;
                _helpFile = string.Empty;
            }
        }

        public void Raise(int number)
        {
            Raise(number, string.Empty);
        }

        public void Raise(int number, string source)
        {
            Raise(number, source, string.Empty);
        }

        public void Raise(int number, string source, string description)
        {
            Raise(number, source, description, string.Empty);
        }

        public void Raise(int number, string source, string description, string helpFile)
        {
            Raise(number, source, description, helpFile, string.Empty);
        }

        public void Raise(int number, string source, string description, string helpFile, string helpContext)
        {
            throw new VBScriptRuntimeException(number, source, description, helpFile, helpContext);
        }

        public override string ToString()
        {
            return _number.ToString();
        }

        internal void internalRaise(int number, string source, string description, string helpFile, string helpContext)
        {
            this._number = number;
            this._description = description;
            this._source = source;
            this._helpContext = helpContext;
            this._helpFile = helpFile;
        }

        internal void internalRaise(Exception ex)
        {
            VBScriptRuntimeException vbEx = ex as VBScriptRuntimeException;
            if (vbEx == null)
            {
                vbEx = new VBScriptRuntimeException(ex);
            }

            internalRaise(vbEx.Number, vbEx.Source, vbEx.Description, vbEx.HelpFile, vbEx.HelpContext);
        }

    }
}
