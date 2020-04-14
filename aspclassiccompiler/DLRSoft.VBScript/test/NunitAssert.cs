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

        [ComVisible(false)]
        public void AreEqual(object expected, object actual)
        {
            Assert.AreEqual(expected, actual);
        }

        [ComVisible(false)]
        public void AreEqual(object expected, object actual, string message)
        {
            Assert.AreEqual(expected, actual, message);
        }

        [ComVisible(true)]
        public void AreEqual(object expected, object actual, [Optional]object message)
        {
            if (message == null || message == System.Type.Missing)
                AreEqual(expected, actual);
            else
                AreEqual(expected, actual, message.ToString());
        }

        [ComVisible(false)]
        public void AreNotEqual(object notExpected, object actual)
        {
            Assert.AreNotEqual(notExpected, actual);
        }

        [ComVisible(false)]
        public void AreNotEqual(object notExpected, object actual, string message)
        {
            Assert.AreNotEqual(notExpected, actual, message);
        }

        [ComVisible(true)]
        public void AreNotEqual(object notExpected, object actual, [Optional]object message)
        {
            if (message == null || message == System.Type.Missing)
                AreNotEqual(notExpected, actual);
            else
                AreNotEqual(notExpected, actual, message.ToString());
        }

        [ComVisible(false)]
        public void Fail()
        {
            Assert.Fail();
        }

        [ComVisible(false)]
        public void Fail(string message)
        {
            Assert.Fail(message);
        }

        [ComVisible(true)]
        public void Fail([Optional]object message)
        {
            if (message == null || message == System.Type.Missing)
                Fail();
            else
                Fail(message.ToString());
        }

        [ComVisible(false)]
        public void IsFalse(bool condition)
        {
            Assert.IsFalse(condition);
        }

        [ComVisible(false)]
        public void IsFalse(bool condition, string message)
        {
            Assert.IsFalse(condition, message);
        }

        [ComVisible(true)]
        public void IsFalse(bool condition, [Optional]object message)
        {
            if (message == null || message == System.Type.Missing)
                IsFalse(condition);
            else
                IsFalse(condition, message.ToString());
        }

        [ComVisible(false)]
        public void IsTrue(bool condition)
        {
            Assert.IsTrue(condition);
        }

        [ComVisible(false)]
        public void IsTrue(bool condition, string message)
        {
            Assert.IsTrue(condition, message);
        }

        [ComVisible(true)]
        public void IsTrue(bool condition, [Optional]object message)
        {
            if (message == null || message == System.Type.Missing)
                IsTrue(condition);
            else
                IsTrue(condition, message.ToString());
        }

        #endregion
    }
}
