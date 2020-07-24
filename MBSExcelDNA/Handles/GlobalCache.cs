// ----------------------------------------------------------------------
// IMPORTANT DISCLAIMER:
// The code is for demonstration purposes only, it comes with NO WARRANTY AND GUARANTEE.
// No liability is accepted by the Author with respect any kind of damage caused by any use
// of the code under any circumstances.

// Originally written by Alex Chirokov in https://www.codeproject.com/Articles/1097174/Interpolation-in-Excel-using-Excel-DNA
// Amended by Domenico Picone on 21 07 2020
// ------------------------------------------------------------------------

using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;

//using ExcelDna.Integration;

namespace MBSExcelDNA.Handles
{
    class GlobalCache
    {
        private static HandleStorage m_storage = new HandleStorage();

        internal static object CreateHandle(string objectType, object[] parameters, Func<string, object[], object> maker)
        {
            return m_storage.CreateHandle(objectType, parameters, maker); 
        }

        internal static bool TryGetObject<T>(string name, out T value) where T : class
        {
            return m_storage.TryGetObject<T>(name, out value);
        }

    }
}
