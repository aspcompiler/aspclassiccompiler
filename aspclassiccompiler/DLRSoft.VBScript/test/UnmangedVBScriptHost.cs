using System;
using System.Collections.Generic;
using System.Text;
using MSScriptControl;
using System.IO;
using Dlrsoft.VBScript.Runtime;

namespace Dlrsoft.VBScriptTest
{
    public class UnmangedVBScriptHost
    {
        public static void Run(string filePath)
        {
            ScriptControlClass scripter = new ScriptControlClass();

            try {
                scripter.Language = "VBScript";
                scripter.AllowUI = false;
                scripter.UseSafeSubset = true;
                IAssert assert = new NunitAssert();
                scripter.AddObject("Assert", assert, false);

                string code = File.ReadAllText(filePath);
                scripter.AddCode(code);
            }
            catch(Exception ex)
            {
                throw;
            } 
            finally
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(scripter);
                scripter = null;
            }

        }
    }
}
