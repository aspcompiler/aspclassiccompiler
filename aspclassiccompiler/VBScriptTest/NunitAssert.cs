using System;
using System.Collections.Generic;
using System.Text;
using Dlrsoft.VBScript.Runtime;
using NUnit.Framework;

namespace Dlrsoft.VBScriptTest
{
    public class NunitAssert : IAssert
    {
        #region IAssert Members

        public void AreEqual(object expected, object actual)
        {
            Assert.AreEqual(expected, actual);
        }

        public void AreEqual(object expected, object actual, string message)
        {
            Assert.AreEqual(expected, actual, message);
        }

        public void AreNotEqual(object notExpected, object actual)
        {
            Assert.AreNotEqual(notExpected, actual);
        }

        public void AreNotEqual(object notExpected, object actual, string message)
        {
            Assert.AreNotEqual(notExpected, actual, message);
        }

        public void Fail()
        {
            Assert.Fail();
        }

        public void Fail(string message)
        {
            Assert.Fail(message);
        }

        public void IsFalse(bool condition)
        {
            Assert.IsFalse(condition);
        }

        public void IsFalse(bool condition, string message)
        {
            Assert.IsFalse(condition, message);
        }

        public void IsTrue(bool condition)
        {
            Assert.IsTrue(condition);
        }

        public void IsTrue(bool condition, string message)
        {
            Assert.IsTrue(condition, message);
        }

        #endregion
    }
}
