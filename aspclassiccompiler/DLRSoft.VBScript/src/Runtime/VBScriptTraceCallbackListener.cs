using System;
using System.Collections.Generic;
using System.Text;
using Microsoft.Scripting;
using Microsoft.Scripting.Debugging;
using System.Diagnostics;

namespace Dlrsoft.VBScript.Runtime
{
    public class VBScriptTraceCallbackListener : ITraceCallback
    {
        #region ITraceCallback Members

        public void OnTraceEvent(TraceEventKind kind, string name, string sourceFileName, Microsoft.Scripting.SourceSpan sourceSpan, Microsoft.Scripting.Utils.Func<Microsoft.Scripting.IAttributesCollection> scopeCallback, object payload, object customPayload)
        {
            switch (kind)
            {
                case TraceEventKind.TracePoint:
                case TraceEventKind.FrameEnter:
                case TraceEventKind.FrameExit:
                    Trace.TraceInformation("{4} at {0} Line {1} column {2}-{3}", sourceFileName, sourceSpan.Start.Line, sourceSpan.Start.Column, sourceSpan.End.Column, kind);
                    if (scopeCallback != null)
                    {
                        IAttributesCollection attr = scopeCallback();
                        if (attr != null)
                        {
                            Trace.TraceInformation("Attrs {0}", attr);
                        }
                    }
                    break;
                case TraceEventKind.Exception:
                case TraceEventKind.ExceptionUnwind:
                    //Don't know what to do
                    break;
                case TraceEventKind.ThreadExit:
                    Trace.TraceInformation("Page completed successfully.");
                    break;
                default:
                    //Do nothing
                    break;
            }
        }

        #endregion
    }
}
