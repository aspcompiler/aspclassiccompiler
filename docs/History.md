* 3/29/2011 - Open the source code. See [Announcements](http://weblogs.asp.net/lichen/archive/2011/03/30/asp-classic-compiler-is-now-open-source.aspx).
* Change Set 41463 - 4/4/2010 Version 0.6.2.34834  VS2010 and VS2008 Builds.
# Fixed the bugs regarding unary operation binder and logical binary operation binder reported in this [post](http://aspclassiccompiler.codeplex.com/Thread/View.aspx?ThreadId=205235). 
# Runtime exception thrown by compiled VBScript can now report the name of the source file and line number if you enable tracing. See this [post](http://weblogs.asp.net/lichen/archive/2010/04/05/how-to-enable-tracing-in-asp-classic-compiler.aspx) on how to enable tracing.
* Change Set 41199 - 3/27/2010 Version 0.6.1.34834  VS2010 and VS2008 Builds.
# Fixed worked item 3954.
# Fixed the bugs reported in this [post](http://aspclassiccompiler.codeplex.com/Thread/View.aspx?ThreadId=205235). 
# You can now using custom classes in Asp Classic Compiler. See this [post](http://weblogs.asp.net/lichen/archive/2010/03/28/using-custom-net-classes-in-asp-classic-compiler.aspx) for more details.
* Change Set 33532 - 12/18/2009. Version 0.6.0.34834 VS2010 and VS2008 Builds. (Do not download 33384-33386 as they have build errors)
# We reached a milestone today. ASP Classic Compiler is working now with a first end-to-end sample application, [Fitch and Mather Stocks 1.0](http://www.microsoft.com/downloads/details.aspx?FamilyID=243abaa0-dd03-4fbc-b58f-5da61839f948&DisplayLang=en).
# We have included a [preconfigured copy](History_fmstocks10.zip) of Fitch and Mather application for you to try.
# See [this](http://weblogs.asp.net/lichen/archive/2009/12/20/asp-classic-compiler-0-6-0-released-it-can-run-fmstocks-1-0.aspx) and [this](http://weblogs.asp.net/lichen/archive/2009/12/20/installing-fmstocks-1-0-on-later-platforms.aspx) blog entries for more information.
* Change Set 32739 - 12/3/2009. Version 0.5.7.34834 VS2010 and VS 2008 Builds. 
# Fixed the bug in [this discussion](http://aspclassiccompiler.codeplex.com/Thread/View.aspx?ThreadId=76818).
* Change Set 32557 - 11/30/2009. Version 0.5.6.34834 VS2010 and VS 2008 Builds. Still no Silverlight update.
# Fixed numerous bugs in the parser and the code generator.
# Implemented On Error.
# Allows method redefinition. See this [discussion](http://aspclassiccompiler.codeplex.com/Thread/View.aspx?ThreadId=76620) for more details.
* Change Set 32030 - 11/16/2009. Version 0.5.5.34834 VS2010 and VS2008 Builds. This release does not have a Silverlight build.
# Setup the infrastructure to map from generated code back to the source code so that we can report the location accurately in ASP files.
# If a file has multiple syntax errors, chances are that we are able to report all the errors at once. In contrast, ASP Classic only reports one error at a time. That is because our parser attempts to recover from the error and continue parsing the file. 
* Change Set 31328 - 11/3/2009. Silverlight samples uploaded. Minor bug fix to ASP include syntax. Visit [this blog](http://weblogs.asp.net/lichen/archive/2009/11/04/uploaded-silverlight-sample-for-vbscript-net-compiler.aspx) for details of the Silverlight sample.
* Change Set 31281 - 11/2/2009. Version 0.5.4.34834. Built with DLR 0.92. We are going to stay with DLR 0.92 stable build for a while instead of using the weekly build so that we can be compatible with VS2010 beta2.
# VS2010 build is available now in addition to the VS2008 build.
# Silverlight build of our VBScript.NET compiler is available for the first time. We will try to come out with samples in the next weekly release.
# All the binaries are consolidated into the bin folder.

* Change Set 30945 - 10/25/2009. Version 0.5.3.32871. Built with DLR Changeset 32871.
# Now ASP Classic Compiler would run under medium trust or Windows Azure partial-trust. Partial-trust is the default trust level of Windows Azure. Windows Azure does not support ASP Classic. Windows Azure supports php but requires full-trust. ASP Classic Compiler is the only way to run ASP pages on Azure without changing the trust level.
# Azure sample uploaded. We modified only one view /home/index.asp but that is enough to prove the concept.
# Variable initialization in declaration statement as a VBScript extension. For example: Dim s = New System.Text.StringBuilder() 
# Add the operator overloading through extension methods. C# supports extension methods, but not "extension operators". We added this support so that we can add new behavior without the need to modify our binders. Some typical ASP code would now run much faster by a simple change. For example, ASP pages often have code like:
{{
<%
    Dim s
    Dim i
    
    s = "<table>"
    For i = 1 To 12
        s = s + "<tr>"
        s = s + "<td>" + i + "</td>"
        s = s + "<td>" + MonthName(i) + "</td>"
        s = s + "</tr>"
    Next
    s = s + "</table>"
    Response.Write(s)

%>
}}
This code is slow because the runtime keeps creating new strings and discarding the old one. Now we just need a simple change:
{{
<%
    Imports System
    Dim s = New System.Text.StringBuilder()
    Dim i
    
    s = s + "<table>"
    For i = 1 To 12
        s = s + "<tr>"
        s = s + "<td>" + i + "</td>"
        s = s + "<td>" + MonthName(i) + "</td>"
        s = s + "</tr>"
    Next
    s = s + "</table>"
    Response.Write(s)

%>
}}
Note that we added "+" operator as an extension operator for the class System.Text.StringBuilder.

* Change Set 30470 - 10/17/2009. Version 0.5.2.32072. Built with DLR Changeset 32072
# Fixed a bug so that the default properties are retrieved when concatenating two COM objects.
# Added support for extension methods. This is crucial for ASP.NET MVC support since many HtmlHelper methods are extension methods.
# Created ASP view engine for ASP.NET MVC.
# Ported [NerdDinner sample](http://nerddinner.codeplex.com) to use the ASP view engine. See [Port NerdDinner Sample to ASP View Engine](Port-NerdDinner-Sample-to-ASP-View-Engine) for more details.
* Change Set 29666 - 10/08/2009*  Initial release of version 0.5.1.31188. Built with DLR Changeset 31188.