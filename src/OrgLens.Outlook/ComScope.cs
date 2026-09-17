using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;

namespace OrgLens.Outlook
{
    internal sealed class ComScope : IDisposable
    {
        private readonly Stack<object> objects = new Stack<object>();

        public T Own<T>(T value) where T : class
        {
            if (value != null) objects.Push(value);
            return value;
        }

        public void Dispose()
        {
            while (objects.Count > 0)
            {
                var value = objects.Pop();
                if (Marshal.IsComObject(value)) Marshal.ReleaseComObject(value);
            }
        }
    }
}
