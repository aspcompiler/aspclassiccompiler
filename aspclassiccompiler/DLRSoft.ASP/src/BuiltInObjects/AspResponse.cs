using System;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Web;
using ASPTypeLibrary;

namespace Dlrsoft.Asp.BuiltInObjects
{
	/// <summary>
	/// The response object that is accessible from ASP code
    /// Get not implement IResponse because the COM interface not supported by C#
	/// </summary>
	public class AspResponse //: IResponse
	{
		public HttpContext _context;

		public AspResponse()
		{
			_context=HttpContext.Current;
		}

		public AspResponse(HttpContext context)
		{_context=context;}

        #region IResponse Members

        public void Add(string bstrHeaderValue, string bstrHeaderName)
        {
            AddHeader(bstrHeaderName, bstrHeaderValue);
        }

        public void AddHeader(string name, string value)
        { _context.Response.AddHeader(name, value); }

        public void AppendToLog(string param)
        { _context.Response.AppendToLog(param); }

        public void BinaryWrite(object varInput)
        {
            _context.Response.BinaryWrite((byte[])varInput);
        }

        public void Clear()
        { _context.Response.Clear(); }

        public void End()
        { _context.Response.End(); }

        public void Flush()
        { _context.Response.Flush(); }

        public bool IsClientConnected()
        {
            return _context.Response.IsClientConnected;
        }

        public void Pics(string value)
        { _context.Response.Pics(value); }

        public void Redirect(string url)
        {
            _context.Response.Redirect(url, true);
        }

        public void Write(object output)
        {
            if (output == null) return;

            string strOut = output as string;
            if (strOut == null)
            {
                Type t = output.GetType();
                while (t.IsCOMObject)
                {
                    output = t.InvokeMember(string.Empty,
                        BindingFlags.Default | BindingFlags.Public | BindingFlags.IgnoreCase | BindingFlags.Instance | BindingFlags.GetProperty,
                        null,
                        output,
                        null);
                    t = output.GetType();
                }

                strOut = output.ToString();
            }
            _context.Response.Write(strOut);
        }

        public void WriteBlock(short iBlockNumber)
        {
            throw new NotImplementedException();
        }

        public bool Buffer
        {
            get { return _context.Response.Buffer; }
            set { _context.Response.Buffer = value; }
        }

        public string CacheControl
        {
            get { return _context.Response.CacheControl; }
            set { _context.Response.CacheControl = value; }
        }

        public string CharSet
        {
            get { return _context.Response.Charset; }
            set { _context.Response.Charset = value; }
        }

        public int CodePage
        {
            get
            {
                throw new NotImplementedException();
            }
            set
            {
                throw new NotImplementedException();
            }
        }

        public string ContentType
        {
            get { return _context.Response.ContentType; }
            set { _context.Response.ContentType = value; }
        }

        public IRequestDictionary Cookies
        {
            get
            {
                return new AspCookieCollection(_context.Response.Cookies, false);
            }
        }

        public int Expires
        {
            get { return _context.Response.Expires; }
            set { _context.Response.Expires = value; }
        }

        public DateTime ExpiresAbsolute
        {
            get { return _context.Response.ExpiresAbsolute; }
            set { _context.Response.ExpiresAbsolute = value; }
        }

        public int LCID
        {
            get
            {
                throw new NotImplementedException();
            }
            set
            {
                throw new NotImplementedException();
            }
        }

        public string Status
        {
            get { return _context.Response.Status; }
            set { _context.Response.Status = value; }
        }

        #endregion
    }
}
