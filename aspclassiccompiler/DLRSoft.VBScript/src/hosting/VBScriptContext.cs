using System;
using System.Collections.Generic;
using System.Text;
using Microsoft.Scripting.Runtime;

using System.Dynamic;
#if USE35
// Needed for type language implementers need to support DLR Hosting, such as
// ScriptCode.
using Microsoft.Scripting.Ast;
#else
using System.Linq.Expressions;
#endif
using Microsoft.Scripting;
using Dlrsoft.VBScript.Runtime;
using Dlrsoft.VBScript.Compiler;
using Microsoft.Scripting.Utils;

namespace Dlrsoft.VBScript.Hosting
{
    public class VBScriptContext : LanguageContext
    {
        private readonly VBScript _vbscript;

        public VBScriptContext(ScriptDomainManager manager,
                                IDictionary<string, object> options)
            : base(manager) {
            // TODO: parse options
            // TODO: register event  manager.AssemblyLoaded
            _vbscript = new VBScript(manager.GetLoadedAssemblyList(), manager.Globals);

            if (options.ContainsKey("Trace") && options["Trace"].Equals(true))
            {
                _vbscript.Trace = true;
            }
        }

        public override ScriptCode CompileSourceCode(SourceUnit sourceUnit, CompilerOptions options, ErrorSink errorSink)
        {
            using (var reader = sourceUnit.GetReader()) {
                try
                {
                    switch (sourceUnit.Kind)
                    {
                        case SourceCodeKind.SingleStatement:
                        case SourceCodeKind.Expression:
                        case SourceCodeKind.AutoDetect:
                        case SourceCodeKind.InteractiveCode:
                            return new VBScriptCode(
                                _vbscript, _vbscript.ParseExprToLambda(reader),
                                sourceUnit);
                        case SourceCodeKind.Statements:
                        case SourceCodeKind.File:
                            return new VBScriptCode(
                                _vbscript,
                                _vbscript.ParseFileToLambda(sourceUnit.Path, reader),
                                sourceUnit);
                        default:
                            throw Assert.Unreachable;
                    }
                }
                catch (VBScriptCompilerException ex)
                {
                    VBScriptSourceCodeReader vbscriptReader = reader as VBScriptSourceCodeReader;
                    if (vbscriptReader != null)
                    {
                        ISourceMapper mapper = vbscriptReader.SourceMapper;
                        if (mapper != null)
                        {
                            foreach (VBScriptSyntaxError error in ex.SyntaxErrors)
                            {
                                DocSpan docSpan = mapper.Map(error.Span);
                                error.FileName = docSpan.Uri;
                                error.Span = docSpan.Span;
                            }
                        }
                    }
                    throw ex;
                }
                catch (Exception e)
                {
                    // Real language implementation would have a specific type
                    // of exception.  Also, they would pass errorSink down into
                    // the parser and add messages while doing tighter error
                    // recovery and continuing to parse.
                    errorSink.Add(sourceUnit, e.Message, SourceSpan.None, 0,
                                  Severity.FatalError);
                    return null;
                }
            }
        }
    }
}
