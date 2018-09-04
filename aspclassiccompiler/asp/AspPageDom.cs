//////////////////////////////////////////////////////////////////////////////
// Created by: Li Chen
// http://www.dotneteer.com/weblog
// dotneteer@gmail.com
// This is an experimental code. Please do not use in production environment.
// April 22, 2009 V0.1
/////////////////////////////////////////////////////////////////////////////
using System;
using System.Collections.Generic;
using System.Text;
using System.Text.RegularExpressions;
using System.Web;
using System.IO;
using Microsoft.Scripting;
using Dlrsoft.VBScript.Compiler;

namespace Dlrsoft.Asp
{
    public class AspPageDom
    {
        private string _pagePath;
        private string _virtualRootPath;
        private List<string> _literals = new List<string>();
        private StringBuilder _sb;
        private VBScriptSourceMapper _mapper = new VBScriptSourceMapper();
        private int _curLine = 0;

        public AspPageDom()
        {
            _sb = new StringBuilder();
            _literals = new List<string>();
        }

        /// <summary>
        /// Read the page from pagePath and process the page.
        /// </summary>
        /// <param name="pagePath"></param>
        public void processPage(string pagePath)
        {
            processPage(pagePath, _virtualRootPath);
        }

        /// <summary>
        /// Read the page from pagePath and process the page.
        /// </summary>
        /// <param name="pagePath"></param>
        /// <param name="virtualRootPath">Physical path of the virtual root used when this class is used
        /// outside ASP.NET so then HttpContext is not available</param>
        public void processPage(string pagePath, string virtualRootPath)
        {
            if (string.IsNullOrEmpty(this._pagePath))
            {
                this._pagePath = pagePath;
            }
            string aspFile = File.ReadAllText(pagePath);
            processPage(pagePath, aspFile, virtualRootPath);
        }

        /// <summary>
        /// Process the page in aspFile string. Pagepath is only used for getting included files.
        /// </summary>
        /// <param name="pagePath"></param>
        /// <param name="aspFile"></param>
        /// <param name="virtualRootPath">Physical path of the virtual root used when this class is used
        public void processPage(string pagePath, string aspFile, string virtualRootPath)
        {
            _virtualRootPath = virtualRootPath;
            Range[] lineRanges = SourceUtil.GetLineRanges(aspFile);
            const string pattern = "<%(?<contents>.*?)%>|<!--\\s*#include\\s+(?<contents>.*?)\\s*-->|<script[^>]+runat=\"?server\"?[^>]*>(?<contents>.*?)</script>";
            Regex r = new Regex(pattern, RegexOptions.IgnoreCase | RegexOptions.Singleline);
            MatchCollection ms = r.Matches(aspFile);
            int p1 = 0;
            int p2 = 0;
            foreach (Match m in ms)
            {
                p2 = m.Index;
                if (p2 - p1 > 0)
                    appendBlock(pagePath, SourceUtil.GetSpan(lineRanges, p1, p2 - 1), GetListeral(aspFile, p1, p2), 1);
                p1 = m.Index + m.Length;
                string value = m.Value;
                string contents = m.Groups["contents"].Value;
                switch (value.Substring(0, 3))
                {
                    case "<!-": //Include
                        processInclude(pagePath, contents.ToLower());
                        break;
                    case "<%@": //Declaration. Ignore
                        break;
                    default:
                        string temp = contents.Trim();
                        if (!string.IsNullOrEmpty(temp))
                        {
                            if (temp[0] == '=') //Expression
                            {
                                contents = string.Format("response.Write({0})", temp.Substring(1).Trim());
                                appendBlock(pagePath, SourceUtil.GetSpan(lineRanges, m.Index, p1 - 1), contents, 1);
                            }
                            else
                            {
                                appendBlock(pagePath, SourceUtil.GetSpan(lineRanges, m.Index, p1 - 1), contents, SourceUtil.GetLineCount(contents));
                            }
                        }
                        break;
                }
            }
            p2 = aspFile.Length;
            if (p2 - p1 > 0)
                appendBlock(pagePath, SourceUtil.GetSpan(lineRanges, p1, p2 - 1), GetListeral(aspFile, p1, p2), 1);
        }

        private string GetListeral(string aspFile, int p1, int p2)
        {
            _literals.Add(aspFile.Substring(p1, p2 - p1));
            return string.Format("response.Write(literals({0}))", _literals.Count - 1);
        }

        private void processInclude(string parent, string spec)
        {
            string filePath = null;
            int x = spec.IndexOf('=');
            string attrib = spec.Substring(0, x).Trim();
            string value = spec.Substring(x + 1).Trim();
            if (value.StartsWith("\""))
            {
                value = value.Substring(1, value.Length - 2);
            }
            switch(attrib)
            {
                case "file":
                    if (Path.IsPathRooted(value))
                        filePath = value;
                    else
                        filePath = Path.Combine(Path.GetDirectoryName(parent), value);
                    break;
                case "virtual":
                    if (HttpContext.Current != null)
                    {
                        filePath = HttpContext.Current.Server.MapPath(value);
                    }
                    else if (!string.IsNullOrEmpty(_virtualRootPath))
                    {
                        filePath = Path.Combine(_virtualRootPath, value);
                    }
                    else
                    {
                        throw new ArgumentException("Must supply the VirtualRootPath if not running under an HttpContext.");
                    }
                    break;
                default:
                    throw new ArgumentException("Invalid include spec:" + spec);
            }

            processPage(filePath);
        }

        private void appendBlock(string filepath, SourceSpan span, string contents, int lines)
        {
            int startLine = _curLine + 1;
            int startIndex = _sb.Length;
            _curLine += lines;
            int endLine = _curLine;
            _sb.AppendLine(contents);
            int endIndex = _sb.Length;

            int startColumn = 1;
            int endColumn;
            if (lines == 0)
            {
                endColumn = contents.Length;
            }
            else
            {
                int x = contents.LastIndexOf('\r');
                if (x < contents.Length - 1 && contents[x + 1] == '\n')
                {
                    x++;
                }
                endColumn = contents.Length - x;
            }

            SourceSpan generatedSpan = new SourceSpan(
                new SourceLocation(startIndex, startLine, startColumn),
                new SourceLocation(endIndex, endLine, endColumn)
            );
            _mapper.AddMapping(generatedSpan, new DocSpan(filepath, span));
        }

        public string PagePath
        {
            get { return _pagePath; }
        }

        public string Code
        {
            get { return _sb.ToString(); }
        }

        public List<string> Literals
        {
            get { return _literals; }
        }

        public VBScriptSourceMapper Mapper
        {
            get { return _mapper; }
        }
    }
}
