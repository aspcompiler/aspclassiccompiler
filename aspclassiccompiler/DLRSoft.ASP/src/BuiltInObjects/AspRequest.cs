using System;
using System.Web;
using System.Collections.Generic;
using System.Text;
using ASPTypeLibrary;

namespace Dlrsoft.Asp.BuiltInObjects
{
    public class AspRequest : IRequest
    {
        private HttpContext _context;

        #region constructor
		public AspRequest()
		{
			_context= HttpContext.Current;
		}

		public AspRequest(HttpContext context)
		{
			_context=context;
		}
        #endregion

        #region IRequest Members

        public object BinaryRead(ref object pvarCountToRead)
        {
            return _context.Request.BinaryRead(Convert.ToInt32(pvarCountToRead));
        }

        public IRequestDictionary Body
        {
            get { throw new NotImplementedException(); }
        }

        public IRequestDictionary ClientCertificate
        {
            get { throw new NotImplementedException(); }
        }

        public IRequestDictionary Cookies
        {
            get { return new AspCookieCollection(_context.Request.Cookies, true); }
        }

        public IRequestDictionary Form
        {
            get { return new AspNameValueCollection(_context.Request.Form); }
        }

        public IRequestDictionary QueryString
        {
            get { return new AspNameValueCollection(_context.Request.QueryString); }
        }

        public IRequestDictionary ServerVariables
        {
            get { return new AspNameValueCollection(_context.Request.ServerVariables); }
        }

        public int TotalBytes
        {
            get { return _context.Request.TotalBytes; }
        }

        public object this[string key]
        {
            get
            {
                if (_context.Request.Form[key] != null)
                    return _context.Request.Form[key];
                else if (_context.Request.QueryString[key] != null)
                    return _context.Request.QueryString[key];
                else
                    return "";
            }
        }

        #endregion
    }
}
