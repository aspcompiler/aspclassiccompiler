using System;
using System.Collections.Generic;
using System.Text;
using System.Dynamic;
using Microsoft.Scripting.Hosting;
#if USE35
using Microsoft.Scripting.Ast;
#else
using System.Linq.Expressions;
#endif

using Dlrsoft.VBScript.Hosting;
namespace Dlrsoft.VBScript
{
    class Program
    {
        static void Main(string[] args)
        {
            // Setup DLR ScriptRuntime with our languages.  We hardcode them here
            // but a .NET app looking for general language scripting would use
            // an app.config file and ScriptRuntime.CreateFromConfiguration.
            var setup = new ScriptRuntimeSetup();
            string qualifiedname = typeof(VBScriptContext).AssemblyQualifiedName;
            setup.LanguageSetups.Add(new LanguageSetup(
                qualifiedname, "vbscript", new[] { "vbscript" }, new[] { ".vbs" }));
            var dlrRuntime = new ScriptRuntime(setup);
            // Don't need to tell the DLR about the assemblies we want to be
            // available, which the SymplLangContext constructor passes to the
            // Sympl constructor, because the DLR loads mscorlib and System by
            // default.
            //dlrRuntime.LoadAssembly(typeof(object).Assembly);
            dlrRuntime.LoadAssembly(typeof(global::Dlrsoft.VBScript.Runtime.BuiltInFunctions).Assembly);

            // Get a Sympl engine and run stuff ...
            var engine = dlrRuntime.GetEngine("vbscript");
            //string filename = @"..\..\test\test.vbs";
            string filename = args[0];
            var scriptSource = engine.CreateScriptSourceFromFile(filename);
            var compiledCode = scriptSource.Compile();
            var feo = engine.CreateScope(); //File Level Expando Object
            //feo = engine.ExecuteFile(filename, feo);
            feo.SetVariable("response", System.Console.Out);
            compiledCode.Execute(feo);
            Console.WriteLine("Type any key to continue..");
            Console.Read();
        }
    }
}
