//////////////////////////////////////////////////////////////////////////////
// Created by: Li Chen
// http://www.dotneteer.com/weblog
// dotneteer@gmail.com
// This is an experimental code. Please do not use in production environment.
// April 22, 2009 V0.1
/////////////////////////////////////////////////////////////////////////////
using System;
using System.Collections.Generic;
using System.Text;
using Microsoft.Scripting.Hosting;

namespace Dlrsoft.Asp
{
    public class CompiledPage
    {
        private CompiledCode _code;
        private List<string> _literals;
        private DateTime _compileTime;

        public CompiledPage(CompiledCode code, List<string> literals)
        {
            this._code = code;
            this._literals = literals;
            _compileTime = DateTime.Now;
        }

        public CompiledCode Code
        {
            get { return _code; }
        }

        public List<string> Literals
        {
            get { return _literals; }
        }

        public DateTime CompileTime
        {
            get { return _compileTime; }
        }
    }
}
