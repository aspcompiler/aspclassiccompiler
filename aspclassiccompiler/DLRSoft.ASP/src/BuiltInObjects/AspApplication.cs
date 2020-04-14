using System;
using System.Collections;
using System.Collections.Specialized;
using System.Web;
using ASPTypeLibrary;

namespace Dlrsoft.Asp.BuiltInObjects
{
	/// <summary>
	/// The application object that is accessible from ASP code
	/// </summary>
	public class AspApplication : IApplicationObject
	{
		private HttpContext _context;
		private bool _locked = false;

		public AspApplication()
		{
			_context=HttpContext.Current;
		}

		public AspApplication(HttpContext context)
		{
			_context=context;
		}

        #region IApplicationObject Members

        public IVariantDictionary Contents
        {
            get { return new AspVariantDictionary(_context.Application.Contents); }
        }

        public IVariantDictionary StaticObjects
        {
            get { return new AspStaticObjectsVariantDictionary(_context.Application.StaticObjects); }
        }

        public bool Locked
        {
            get { return _locked; }
        }

        public void Lock()
        {
            _context.Application.Lock();
            _locked = true;
        }

        public void UnLock()
        {
            _context.Application.UnLock();
            _locked = false;
        }

        public object this[string key]
        {
            get { return _context.Application[key]; }
            set { _context.Application[key] = value; }
        }

        public void let_Value(string bstrValue, object pvar)
        {
            this[bstrValue] = pvar;
        }

        #endregion
    }
}
