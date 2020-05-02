using System;
using System.Collections.Generic;
using System.Text;
using NUnit.Framework;
using System.IO;

namespace Dlrsoft.VBScriptTest
{
    [TestFixture]
    public class UnmanagedVBScriptTest
    {
        [Test]
        public void RunTest()
        {
            DirectoryWalker.Walk(
                Path.Combine(DirectoryWalker.AssemblyDirectory, "../../VBScripts"),
                f =>
                {
                    if (f.EndsWith(".vbs", StringComparison.InvariantCultureIgnoreCase))
                    {
                        try
                        {
                            UnmangedVBScriptHost.Run(f);
                        }
                        catch (Exception ex)
                        {
                            Assert.Fail(f + ":" + ex.Message);
                        }
                    }
                }
            );
        }
    }
}
