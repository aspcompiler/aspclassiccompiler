using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using ASPTypeLibrary;

namespace Dlrsoft.Asp.BuiltInObjects
{
    internal class AspStringList : IStringList
    {
        private String[] _values;

        public AspStringList(string[] values)
        {
            this._values = values;
        }

        public override string ToString()
        {
            return string.Join(",", _values);
        }

        #region IStringList Members

        public IEnumerator GetEnumerator()
        {
            return _values.GetEnumerator();
        }

        public int Count
        {
            get { return _values.Length; }
        }

        public object this[object i]
        {
            get { return _values[(int)i - 1]; }
        }

        #endregion
    }
}
