using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.IO;
using System.Web;
using Dlrsoft.VBScript.Compiler;

namespace Dlrsoft.Asp
{
    public class AspHelper
    {
        public static void RenderError(HttpResponse response, VBScriptCompilerException exception)
        {
            response.Clear();
            response.StatusCode = 500;
            response.StatusDescription = "Internal Server Error";

            RenderError(response.Output, exception);
            response.End();
        }

        public static void RenderError(TextWriter output, VBScriptCompilerException exception)
        {
            output.Write("<h1>VBScript Compiler Error</h1>");
            output.Write("<table>");
            output.Write("<tr>");
            output.Write(string.Format("<td>{0}</td>", "FileName"));
            output.Write(string.Format("<td>{0}</td>", "Line"));
            output.Write(string.Format("<td>{0}</td>", "Column"));
            output.Write(string.Format("<td>{0}</td>", "Error Code"));
            output.Write(string.Format("<td>{0}</td>", "Error Description"));
            output.Write("</tr>");
            foreach (VBScriptSyntaxError error in exception.SyntaxErrors)
            {
                output.Write("<tr>");
                output.Write(string.Format("<td>{0}</td>", error.FileName));
                output.Write(string.Format("<td>{0}</td>", error.Span.Start.Line));
                output.Write(string.Format("<td>{0}</td>", error.Span.Start.Column));
                output.Write(string.Format("<td>{0}</td>", error.ErrorCode));
                output.Write(string.Format("<td>{0}</td>", error.ErrorDescription));
                output.Write("</tr>");
            }
            output.Write("</table>");
        }
    }
}
