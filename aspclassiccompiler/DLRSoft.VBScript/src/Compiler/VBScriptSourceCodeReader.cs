using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using Microsoft.Scripting;

namespace Dlrsoft.VBScript.Compiler
{
    public class VBScriptSourceCodeReader : SourceCodeReader
    {
        private ISourceMapper _mapper;

        public VBScriptSourceCodeReader(TextReader textReader, Encoding encoding, ISourceMapper mapper)
            : base(textReader, encoding)
        {
            _mapper = mapper;
        }

        public ISourceMapper SourceMapper
        {
            get { return _mapper; }
        }
    }
}
