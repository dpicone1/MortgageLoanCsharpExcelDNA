// ----------------------------------------------------------------------
// IMPORTANT DISCLAIMER:
// The code is for demonstration purposes only, it comes with NO WARRANTY AND GUARANTEE.
// No liability is accepted by the Author with respect any kind of damage caused by any use
// of the code under any circumstances.
// Any market parameters used are not real data but have been created to clarify the exercises 
// and should not be viewed as actual market data.
//
// Author Domenico Picone 
// ------------------------------------------------------------------------
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace MBSExcelDNA.Excel
{
    //Casting from C# array To Excel
    public class ExcelCasting
    {
        
        public static object[,] toColumnVect<T>(T[] o)
        {
            object[,] outPut = new object[o.Count(), 1];
            for (int i = 0; i < o.Count(); i++)
            {
                outPut[i, 0] = o[i];
            }
            return outPut;
        }

        public static T[] myArray<T>(Object[,] O)
        {
            if ((O.GetLowerBound(1) == O.GetUpperBound(1)) & (O.GetUpperBound(0) != O.GetUpperBound(1)))
            {
                List<T> l = new List<T>();
                for (int i = O.GetLowerBound(0); i <= O.GetUpperBound(0); i++)
                {
                    int firstCol = O.GetLowerBound(1);
                    l.Add((T)O[i, firstCol]);
                }
                return l.ToArray<T>();
            }

            if ((O.GetLowerBound(0) == O.GetUpperBound(0)) & (O.GetUpperBound(0) != O.GetUpperBound(1)))
            {
                List<T> l = new List<T>();
                for (int i = O.GetLowerBound(1); i <= O.GetUpperBound(1); i++)
                {
                    int firstCol = O.GetLowerBound(0);
                    l.Add((T)O[firstCol, i]);
                }
                return l.ToArray<T>();
            }
            return null;
        }

    }
}
