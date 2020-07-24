using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using System.Windows.Forms;
using ExcelDna.Integration;
using System.Diagnostics;
using ExcelDna.Logging;

using MBSExcelDNA.Handles;
using MBSExcelDNA.Loan;
using MBSExcelDNA.Global;

namespace MBSExcelDNA
{
    public class ExcelLiborRate
    {
        private static readonly object m_sync = new object();
        private static readonly string m_tagLibor = "#LiborRate";

        [ExcelFunction(Description = "Prepare LiborCurve")]
        public static object LiborCurve_create([ExcelArgument(Description = @"Libor Curve as an array")] double[] Libor_Array_)
        {
            if (ExcelDnaUtil.IsInFunctionWizard())
                return ExcelError.ExcelErrorRef;
            else
            {
                // Libor Curve
                int len = Libor_Array_.Length;
                Debug.Assert(len <= GlobalVar.GlobalMaxMortgageLoanMaturity, "Libor Curve should be EXACTLY of 360 data points" + GlobalVar.GlobalMaxMortgageLoanMaturity);

                return GlobalCache.CreateHandle(m_tagLibor, new object[] { Libor_Array_, len, "LiborCurve_create" },
                    (objectType, parameters) =>
                    {
                        LiborRates BoERate_Array = construct_LiborCurve(Libor_Array_);
                        if (BoERate_Array == null)
                            return ExcelError.ExcelErrorNull;
                        else
                            return BoERate_Array;
                    });
            }
        }

        private static LiborRates construct_LiborCurve(double[] curva)
        {
            LiborRates LiborRate_Array = null;

            try
            {
                LiborRate_Array = new LiborRates(curva);
            }
            catch (Exception ex)
            {
                lock (m_sync)
                {
                    LogDisplay.WriteLine("Error: " + ex.ToString());
                }
            }

            return LiborRate_Array;
        }
    }
}
