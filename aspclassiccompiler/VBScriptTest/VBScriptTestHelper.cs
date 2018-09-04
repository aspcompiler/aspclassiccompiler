using System;
using System.Collections.Generic;
using System.Text;
using System.Dynamic;
#if USE35
using Microsoft.Scripting.Hosting;
using Microsoft.Scripting.Ast;
#else
using System.Linq.Expressions;
#endif

using Dlrsoft.VBScript.Hosting;
namespace Dlrsoft.VBScriptTest
{
    public class VBScriptTestHelper
    {
        static public void Run(string filePath)
        {
            // Setup DLR ScriptRuntime with our languages.  We hardcode them here
            // but a .NET app looking for general language scripting would use
            // an app.config file and ScriptRuntime.CreateFromConfiguration.
            var setup = new ScriptRuntimeSetup();
            string qualifiedname = typeof(VBScriptContext).AssemblyQualifiedName;
            setup.LanguageSetups.Add(new LanguageSetup(
                qualifiedname, "vbscript", new[] { "vbscript" }, new[] { ".vbs" }));
            var dlrRuntime = new ScriptRuntime(setup);

            //Add the VBScript runtime assembly
            dlrRuntime.LoadAssembly(typeof(global::Dlrsoft.VBScript.Runtime.BuiltInFunctions).Assembly);

            // Get a VBScript engine and run stuff ...
            var engine = dlrRuntime.GetEngine("vbscript");
            var scriptSource = engine.CreateScriptSourceFromFile(filePath);
            var compiledCode = scriptSource.Compile();
            var feo = engine.CreateScope();
            //feo = engine.ExecuteFile(filename, feo);
            feo.SetVariable("Assert", new NunitAssert());
            compiledCode.Execute(feo);
        }
    }
}
