// ----------------------------------------------------------------------
// IMPORTANT DISCLAIMER:
// The code is for demonstration purposes only, it comes with NO WARRANTY AND GUARANTEE.
// No liability is accepted by the Author with respect any kind of damage caused by any use
// of the code under any circumstances.
// Any market parameters used are not real data but have been created to clarify the exercises 
// and should not be viewed as actual market data.
//
// 
// Author Domenico Picone
// ------------------------------------------------------------------------
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Reflection;

namespace MBSExcelDNA.Loan
{
    public class Common
    {
        public static Type[] GetClassTypes(Assembly assembly, string nameSpace)
        {
            return assembly.GetTypes().Where(t => t.IsClass && String.Equals(t.Namespace, nameSpace)).ToArray();
        }
    }
}