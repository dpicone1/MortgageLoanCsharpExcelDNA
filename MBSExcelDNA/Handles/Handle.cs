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

using ExcelDna.Integration;

namespace MBSExcelDNA.Handles
{
    class Handle : IExcelObservable, IDisposable
    {
        private static readonly object m_lock = new object();
        private static int m_index;

        private readonly HandleStorage m_storage;
        private IExcelObserver m_observer;
        private readonly string m_name;
        private readonly object m_value;

        public Handle(HandleStorage storage, string tag, object value)
        {
            m_storage = storage;
            m_value = value;

            lock (m_lock)
            {
                m_name = String.Format("{0}:{1}", tag, m_index++);
            }
        }

        public IDisposable Subscribe(IExcelObserver observer)
        {
            m_observer = observer;
            m_observer.OnNext(m_name);
            return this;
        }

        public void Dispose()
        {
            m_storage.Remove(this); 
        }

        public string Name
        {
            get
            {
                return m_name;
            }
        }

        public object Value
        {
            get 
            {
                return m_value;
            }
        }
    }
}
