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
using System.Threading.Tasks;

namespace MBSExcelDNA.Loan
{
    public interface IRepayment
    {
        double pmt(double Balance, double Rate, int RemainingPeriods);
    }

    public abstract class Repayment : IRepayment
    {
        public abstract double pmt(double Balance, double Rate, int RemainingPeriods);
    }

    public class RepaymentPI : Repayment
    {
        public override double pmt(double Balance, double Rate, int RemainingPeriods)
        {
            double res = 0.0;
            double coupon = Rate / 12.0;
            double disc = System.Math.Pow((1 + coupon), RemainingPeriods);
            if (Balance > 0.0001) { res = Balance * coupon * disc / (disc - 1); }
            else { res = 0; }

            return res;
        }
    }

    public class RepaymentIO : Repayment
    {
        public override double pmt(double Balance, double Rate, int RemainingPeriods)
        {
            double res = 0;
            if (RemainingPeriods == 1)
            {
                res = Balance + Balance * Rate / 12.0;
            }
            return res;
        }
    }
}
