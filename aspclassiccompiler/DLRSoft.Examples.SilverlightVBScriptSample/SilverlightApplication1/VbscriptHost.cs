using System;
using System.Net;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Documents;
using System.Windows.Ink;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Animation;
using System.Windows.Shapes;
using Microsoft.Scripting;
using Microsoft.Scripting.Hosting;
using Microsoft.Scripting.Silverlight;
using Dlrsoft.VBScript.Hosting;

namespace SilverlightApplication1
{
    public class VbscriptHost
    {
        private ScriptRuntime _runtime = null;
        private ScriptEngine _engine;

        public VbscriptHost()
        {
            var setup = new ScriptRuntimeSetup();
            string qualifiedname = typeof(VBScriptContext).AssemblyQualifiedName;
            setup.LanguageSetups.Add(new LanguageSetup(
                qualifiedname, "vbscript", new[] { "vbscript" }, new[] { ".vbs" }));
            setup.HostType = typeof(BrowserScriptHost);
            _runtime = new ScriptRuntime(setup);
            _engine = _runtime.GetEngine("vbscript");

            _runtime.LoadAssembly(typeof(VBScriptContext).Assembly);
        }

        public ScriptScope CreateScope()
        {
            return _engine.CreateScope();
        }

        public CompiledCode Compile(string code)
        {
            ScriptSource src = _engine.CreateScriptSourceFromString(code, SourceCodeKind.File);
            return src.Compile();
        }
    }
}
