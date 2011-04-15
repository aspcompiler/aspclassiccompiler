using System;
using System.Collections.Generic;
using System.Text;
using Dlrsoft.VBScript.Runtime;
using NUnit.Framework;
using System.Runtime.InteropServices;

namespace Dlrsoft.VBScriptTest
{
    [ComVisible(true)]
    public class NunitAssert : IAssert
    {
        #region IAssert Members

        public void AreEqual(object expected, object actual, string message)
        {
            Assert.AreEqual(expected, actual, message);
        }

        public void AreNotEqual(object notExpected, object actual, string message)
        {
            Assert.AreNotEqual(notExpected, actual, message);
        }

        public void Fail(string message)
        {
            Assert.Fail(message);
        }

        public void IsFalse(bool condition, string message)
        {
            Assert.IsFalse(condition, message);
        }

        public void IsTrue(bool condition, string message)
        {
            Assert.IsTrue(condition, message);
        }

        #endregion
    }
}
