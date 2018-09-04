# Project Description
**This project is no longer actively maintained.**

ASP Classic Compiler contains a VBScript compiler. It compiles ASP classic pages into a .NET Intermediate Language (IL) so that they can run inside the ASP.NET environment. This project allows ASP classic pages to coexist
with ASP.net pages so that ASP Classic site owners can migrate to ASP.net page by page. As there are few ASP classic sites left, this project is no longer active. I had lots of fun working on the project, amongst them:

- The project used ideas like call site caching (also called polymophic inline cache). It is about twice as fast as Microsoft VBScript interpreter.
- I also experimented with type inference and applied it to string concatenation, something often seen in ASP classic code, using StringBuilder. The result is that some pages render hundreds times faster.
- I was awarded Microsoft ASP.NET MVP 4 times from 2010-2014 and got to chance to meet many new people and make new friends.

- 3/30/2014 Converted the source control from Mercurial to Git. Added examples on DateTime usage.
- 4/10/2011 Uploaded unit testing framework. Documents to appear in my blog.
- **3/29/2011. Open the source code under the Apache 2.0 license. See [announcements](http://weblogs.asp.net/lichen/archive/2011/03/30/asp-classic-compiler-is-now-open-source.aspx).**
- Build 0.6.2.34834 uploaded on 04/04/2010. Fixed a few bugs. See [History](docs/History.md) for more details.
- Build 0.6.1.34834 uploaded on 03/27/2010. Fixed a few bugs.
- Build 0.6.0.34834 uploaded on 12/18/2009. It works with FMStocks 1.0 now. See [this blog entry](http://weblogs.asp.net/lichen/archive/2009/12/20/asp-classic-compiler-0-6-0-released-it-can-run-fmstocks-1-0.aspx) for more details
- Build 0.5.7.34834 uploaded on 12/03/2009. Fixed numerous little bugs.
- Build 0.5.6.34834 uploaded on 11/30/2009. Fixed numerous little bugs.
- Build 0.5.5.34834 uploaded on 11/16/2009. This version has significant enhancements on error reporting.
- Build 0.5.4.34834 uploaded on 11/2/2009. Built with DLR 0.92. VS2010 and Silverlight build are available now.
- Build 0.5.3.32871 uploaded on 10/25/2009. This build supports Windows Azure, variable initializer and extension operators.
- Build 0.5.2.32072 uploaded on 10/17/2009. This build supports using asp classic compiler as an ASP.NET MVC ViewEngine.

# Project Detail

ASP Classic Compiler is based on our VBScript compiler (sorry, no JScript at this time) that compiles VBScript into a .NET executable. The compiler host is implemented as an ASP.NET handler. The VBScript compiler is based on a hand-written parser and uses Microsoft [Dynamic Language Runtime](http://www.codeplex.com/dlr) to generate the IL. 

This project has two goals:

- High backward compatibility. The goal is to be able to run the majority of real world ASP pages under ASP.NET unmodified. This will provide a far smoother experience for migrating ASP Classic to ASP.NET. Currently, most ASP projects have to be migrated as a whole because ASP Classic and ASP.NET do not share states. ASP Classic Compiler will run ASP pages under the context of ASP.NET so they share the same state and security. This allows projects to be migrated page by page.
- CLR interoperability. For those who need a scripting language like VBScript, we intend to an a minimum set of new syntax constructs so that it can interoperate with .NET classes without registering these classes as COM.

Please visit the [Features](docs/Features.md) page for features and limitations. If you encounter an incompatibility, do no modify your page yet. Instead, use Issue Tracker to post bugs and feature requests. Take a vote so that we know the priorities. It is our goal to run most real-world ASP pages without modification.

# Typical scenarios to use ASP Classic Compiler

- When you need to run ASP and ASP.NET side-by-site, ASP Classic Compiler runs your pages inside the ASP.NET runtime. You can migrate page by page and take advantage of the ASP.NET infrastructure immediately.
- When you want to run ASP pages faster, ASP Classic Compiler runs the same page up to a few times faster using the latest technology in Dynamic Language Runtime. Pages can run even faster if they are modified to use ASP.NET cache and .NET StringBuilder.
- When you run ASP in Cassini (the built-in web server in ASP.NET), ASP Classic Compiler is the only solution to run ASP pages inside Cassini.
- Windows Azure application in the cloud, as it is an ASP.NET application. Windows Azure does not support ASP Classic. You can run php but you have to run it in full-trust. ASP Classic Compiler can run in standard partial-trust since it is just another ASP application. You can even deploy pages to Azure storage so that you do not need to repackage your application to deploy a change.
- As view in ASP.NET MVC. Since ASP tends to generate html directly, it is more natural to migrate ASP pages to ASP.NET MVC than ASP.NET WebForms. We will provide a view engine so that the ASP Classic Compiler can be used as a view engine for ASP.NET MVC.
- As the template engine for a code generator. The ASP Classic Compiler can be hosted anywhere from a console to a GUI application. We run ASP Classic Compiler in our unit testing project.
- When you want to deploy a dynamic website on a read-only media/devices, it is possible to package ASP Classic Compiler with an embedded web server (e.g. Cassini) and database (e.g., SQL Server Compact edition) to run from a CD-ROM. ASP Classic Compiler generates code in memory. Users can run it when logging into the standard user account without requiring elevated privileges. .NET 2.0 is the only thing that needs to be preinstalled. IIS is not needed.

Have fun!

# Related Pages:
- [Features](docs/Features.md)
- [Installation](docs/Installation.md)
- [FAQ](docs/FAQ.md)
- [History](docs/History.md)
- [ASP View Engine for ASP.NET MVC](docs/Port-NerdDinner-Sample-to-ASP-View-Engine.md)
- [Contact](docs/Contact.md)
- [Acknowledgements](docs/Acknowledgements.md)
- [Contributors](docs/Contributors.md)
