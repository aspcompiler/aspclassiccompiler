using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Web;
using System.Web.Mvc;
using System.IO;
using Microsoft.Scripting;
using Microsoft.Scripting.Runtime;
using Microsoft.Scripting.Hosting;
#if USE35
using Microsoft.Scripting.Generation;
using Microsoft.Scripting.Ast;
#else
using System.Linq.Expressions;
#endif
using Dlrsoft.VBScript.Hosting;
using Dlrsoft.VBScript.Compiler;
using Dlrsoft.Asp.BuiltInObjects;

namespace Dlrsoft.Asp.Mvc
{
    public class AspView : IView
    {
        private static Dictionary<string, CompiledPage> _scriptCache = new Dictionary<string, CompiledPage>();
        private static AspHost aspHost;
        static AspView()
        {
            aspHost = new AspHost(null);
            //aspHost.LoadAssembly(typeof(IView).Assembly);
        }


        public string ViewPath { get; private set; }
        public string MasterPath { get; private set; }

        public AspView(string viewPath)
        {
            this.ViewPath = viewPath;
        }

        public AspView(string viewPath, string masterPath)
		{
			this.ViewPath = viewPath;
			this.MasterPath = masterPath;
		}

        public void Render(ViewContext viewContext, TextWriter writer)
        {
            string pagePath = viewContext.HttpContext.Server.MapPath(this.ViewPath);

            CompiledPage cpage = null;
            if (_scriptCache.ContainsKey(pagePath))
            {
               cpage = _scriptCache[pagePath];
               //don't use it if updated
               if (cpage.CompileTime < File.GetLastWriteTime(pagePath))
                   cpage = null;
            }

            if (cpage == null)
            {
                try
                {
                    cpage = aspHost.ProcessPageFromFile(pagePath);
                    _scriptCache[pagePath] = cpage;
                }
                catch (VBScriptCompilerException ex)
                {
                    AspHelper.RenderError(writer, ex);
                    return;
                }
            }

            ScriptScope responseScope = aspHost.CreateScope();
            HttpContext context = HttpContext.Current;
            responseScope.SetVariable("context", context);
            responseScope.SetVariable("request", context.Request);
            responseScope.SetVariable("session", context.Session);
            responseScope.SetVariable("server", context.Server);
            responseScope.SetVariable("application", context.Application);
            //responseScope.SetVariable("response", new AspResponse(context));

            responseScope.SetVariable("writer", writer);
            responseScope.SetVariable("response", writer);

            responseScope.SetVariable("viewcontext", viewContext);
            responseScope.SetVariable("tempdata", viewContext.TempData);

            ViewDataDictionary viewData = viewContext.ViewData;
            responseScope.SetVariable("viewdata", viewData);
            responseScope.SetVariable("model", viewData.Model);

            ViewDataContainer viewDataContainer = new ViewDataContainer(viewData);
            responseScope.SetVariable("ajax",  new AjaxHelper(viewContext, viewDataContainer));
            responseScope.SetVariable("html", new HtmlHelper(viewContext, viewDataContainer));
            responseScope.SetVariable("url", new UrlHelper(viewContext.RequestContext));

            //responseScope.SetVariable("err", new Microsoft.VisualBasic.ErrObject());
            //Used to get the literals
            responseScope.SetVariable("literals", cpage.Literals);

            try
            {
               object o = cpage.Code.Execute(responseScope);
            }
            catch (Exception ex)
            {
               throw ex;
            }

        }

        private class ViewDataContainer : IViewDataContainer
        {
            public ViewDataContainer(ViewDataDictionary viewData)
            {
                this.ViewData = viewData;
            }

            public ViewDataDictionary ViewData
            {
                get;
                set;
            }
        }
    }
}
