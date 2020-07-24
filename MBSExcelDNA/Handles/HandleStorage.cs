// ----------------------------------------------------------------------
// IMPORTANT DISCLAIMER:
// The code is for demonstration purposes only, it comes with NO WARRANTY AND GUARANTEE.
// No liability is accepted by the Author with respect any kind of damage caused by any use
// of the code under any circumstances.

//
// Originally written by Alex Chirokov in https://www.codeproject.com/Articles/1097174/Interpolation-in-Excel-using-Excel-DNA
// Amended by Domenico Picone on 21 07 2020
// ------------------------------------------------------------------------
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;

using ExcelDna.Integration;

namespace MBSExcelDNA.Handles
{
    class HandleStorage
    {
        private ReaderWriterLockSlim m_lock = new ReaderWriterLockSlim();
        private Dictionary<string, Handle> m_storage = new Dictionary<string, Handle>();

        internal object CreateHandle(string tag, object[] parameters, Func<string, object[], object> maker)
        {
            return ExcelAsyncUtil.Observe(tag, parameters, () =>
            {
                var value = maker(tag, parameters);
                var handle = new Handle(this, tag, value);

                m_lock.EnterWriteLock();

                try
                {
                    m_storage.Add(handle.Name, handle);
                }
                finally
                {
                    m_lock.ExitWriteLock();
                }
                return handle;

            });

        }

        internal bool TryGetObject<T>(string name, out T value)
        {
            bool found = false;

            value = default(T);

            m_lock.EnterReadLock();

            try
            {
                Handle handle;

                if (m_storage.TryGetValue(name, out handle))
                {
                    if (handle.Value is T)
                    {
                        value = (T)handle.Value;
                        found = true;
                    }
                }
            }
            finally
            {
                m_lock.ExitReadLock();
            }
            return found;
        }


        internal void Remove(Handle handle)
        {
            object value;

            if (TryGetObject(handle.Name, out value))
            {
                m_lock.EnterWriteLock();

                try
                {
                    m_storage.Remove(handle.Name);

                    IDisposable disp = value as IDisposable;

                    if (disp != null)
                    {
                        disp.Dispose();
                    }
                }
                finally
                {
                    m_lock.ExitWriteLock();
                }                
            }
        }
    }
}
