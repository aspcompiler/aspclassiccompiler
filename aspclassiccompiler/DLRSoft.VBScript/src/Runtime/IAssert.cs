using System;
using System.Collections.Generic;
using System.Text;

namespace Dlrsoft.VBScript.Runtime
{
    public interface IAssert
    {

        void AreEqual(object expected, object actual);

        void AreEqual(object expected, object actual, string message);

        void AreNotEqual(object notExpected, object actual);

        void AreNotEqual(object notExpected, object actual, string message);

        void Fail();

        void Fail(string message);

        void IsFalse(bool condition);

        void IsFalse(bool condition, string message);

        void IsTrue(bool condition);

        void IsTrue(bool condition, string message);
    }
}
