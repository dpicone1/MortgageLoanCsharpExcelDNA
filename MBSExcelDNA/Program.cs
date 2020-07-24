// ----------------------------------------------------------------------
// IMPORTANT DISCLAIMER:
// The code is for demonstration purposes only, it comes with NO WARRANTY AND GUARANTEE.
// No liability is accepted by the Author with respect any kind of damage caused by any use
// of the code under any circumstances.
//
// 
// Author Domenico Picone
// ------------------------------------------------------------------------
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Diagnostics;

using System.Windows.Forms;
using ExcelDna.Integration;
using ExcelDna.Logging;

using MBSExcelDNA.Handles;
using MBSExcelDNA.Loan;
using MBSExcelDNA.Excel;
using MBSExcelDNA.Global;

namespace MBSExcelDNA
{
    public class MortgagesDNA : IExcelAddIn
    {
        //Excel Auto Open 
        public void AutoOpen() { }
        //Excel Auto Close 
        public void AutoClose(){ }

        public static void RunRisk()
        {
            dynamic xlApp;
            xlApp = ExcelDnaUtil.Application;

            // read loan data
            double Balance     = (double)(xlApp.Range["Balance"].Value2);
            double Rate        = (double)(xlApp.Range["Rate"].Value2);
            double Spread      = (double)(xlApp.Range["Spread"].Value2);
            int    Maturity    = (int)(xlApp.Range["Maturity"].Value2);
            int    Resetting   = (int)(xlApp.Range["ReSetting"].Value2);
            string FixedOrARM  = (string)(xlApp.Range["LoanType"].Value2);
            string PIOrIO      = (string)(xlApp.Range["RepaymentType"].Value2);
            Debug.Assert(Maturity <= GlobalVar.GlobalMaxMortgageLoanMaturity, "Mortgage Loan Maturity should not be greater than " + GlobalVar.GlobalMaxMortgageLoanMaturity);

            // Libor Curve
            double[] Libor_Curve = ExcelCasting.myArray<double>(xlApp.Range["LiborCurve"].Value2);
            int len              = Libor_Curve.Length;
            Debug.Assert(len <= GlobalVar.GlobalMaxMortgageLoanMaturity, "Libor Curve should be EXACTLY of 360 data points" + GlobalVar.GlobalMaxMortgageLoanMaturity);
            LiborRates libor_rates = new LiborRates(Libor_Curve);

            // Build the loan
            IRepayment rmPI    = RepaymentFactory.GetRep(PIOrIO);
            IMortgageLoan loan = MortgageLoanFactory.GetLoan(FixedOrARM, Balance, Maturity, Rate, Resetting, Spread, libor_rates, rmPI);
            loan.CashFlows();
 
            // Write into excel
            xlApp = ExcelDnaUtil.Application;
            try
            {
                dynamic range = xlApp.Range["BegBalanceOutput"];
                range.Value = ExcelCasting.toColumnVect(loan.Write(loan.ReturnBegBalance()));

                range = xlApp.Range["InterestOutput"];
                range.Value = ExcelCasting.toColumnVect(loan.Write(loan.ReturnInterest()));

                range = xlApp.Range["PrincipalOutput"];
                range.Value = ExcelCasting.toColumnVect(loan.Write(loan.ReturnPrincipal()));

                range = xlApp.Range["CashCollectionsOutput"];
                range.Value = ExcelCasting.toColumnVect(loan.Write(loan.ReturnCashCollections()));

                range = xlApp.Range["EndBalanceOutput"];
                range.Value = ExcelCasting.toColumnVect(loan.Write(loan.ReturnEndBalance()));
            }

            catch (Exception e)
            {
                MessageBox.Show("Error:  " + e.ToString());
            }
        }

        [ExcelFunction(Description = "Calculate Loan Amortization")]
        public static object[,] LoanAmortisation(
            [ExcelArgument(Description = @"Balance Notional")] double Balance,
            [ExcelArgument(Description = @"Loan Rate")] double Rate,
            [ExcelArgument(Description = @"Loan Spread, Zero if It is a Fixed Rate Loan")]   double Spread,
            [ExcelArgument(Description = @"Maturity Period")] int Maturity,
            [ExcelArgument(Description = @"Rate Resetting Period, the same as Maturity if it is a Fixed Rate Loan")] int Resetting,
            [ExcelArgument(Description = @"Fixed (Fixed) or ARM (ARM)")] string FixedOrARM,
            [ExcelArgument(Description = @"Principal and Interest (PI) or Interest Only (IO)")] string PIOrIO,
            [ExcelArgument(Description = @"Libor Curve as an array")] double[] Libor_Curve)
        {
            // check Loan Maturity
            Debug.Assert(Maturity <= GlobalVar.GlobalMaxMortgageLoanMaturity, "Mortgage Loan Maturity should not be greater than " + GlobalVar.GlobalMaxMortgageLoanMaturity);

            // Libor Curve
            LiborRates libor_rates = new LiborRates(Libor_Curve);
            int len = Libor_Curve.Length;
            Debug.Assert(len <= GlobalVar.GlobalMaxMortgageLoanMaturity, "Libor Curve should be EXACTLY of 360 data points" + GlobalVar.GlobalMaxMortgageLoanMaturity);

            // Build the loan
            IRepayment rmPI    = RepaymentFactory.GetRep(PIOrIO);
            IMortgageLoan loan = MortgageLoanFactory.GetLoan(FixedOrARM, Balance, Maturity, Rate, Resetting, Spread, libor_rates, rmPI);
            loan.CashFlows();

            double[] BegBal       = loan.Write(loan.ReturnBegBalance());
            double[] Interest     = loan.Write(loan.ReturnInterest());
            double[] Principal    = loan.Write(loan.ReturnPrincipal());
            double[] Collections  = loan.Write(loan.ReturnCashCollections());
            double[] EndBal       = loan.Write(loan.ReturnEndBalance());

            object[,] a = new object[MBSExcelDNA.Global.GlobalVar.GlobalMaxMortgageLoanMaturity, 5];
            for (int i = 0; i < (int)MBSExcelDNA.Global.GlobalVar.GlobalMaxMortgageLoanMaturity; i++)
            {
                a[i, 0] = BegBal[i];
                a[i, 1] = Interest[i];
                a[i, 2] = Principal[i];
                a[i, 3] = Collections[i];
                a[i, 4] = EndBal[i];                
            }
            return a;
        }

        private static readonly object m_sync = new object();
        private static readonly string m_tag = "#MortgageLoan";
        //private static readonly string m_defaultLoan = "Mortgage";
        //private static readonly string m_tagLoanAmortisation = "#acqLoanAmortisation";

        [ExcelFunction(Description = "Compute Collateral Analysis")]
        public static object MortgageLoan_create(
            [ExcelArgument(Description = @"Balance Notional")] double Balance,
            [ExcelArgument(Description = @"Loan Rate")] double Rate,
            [ExcelArgument(Description = @"Loan Spread, Zero if It is a Fixed Rate Loan")]   double Spread,
            [ExcelArgument(Description = @"Maturity Period")] int Maturity,
            [ExcelArgument(Description = @"Rate Resetting Period, the same as Maturity if it is a Fixed Rate Loan")] int Resetting,
            [ExcelArgument(Description = @"Fixed (Fixed) or ARM (ARM)")] string FixedOrARM,
            [ExcelArgument(Description = @"Principal and Interest (PI) or Interest Only (IO)")] string PIOrIO,
            [ExcelArgument(Description = @"Libor Curve as an name")] string LiborCurve)
        {

            if (ExcelDnaUtil.IsInFunctionWizard())
                return ExcelError.ExcelErrorRef;
            else
            {
                return GlobalCache.CreateHandle(m_tag, new object[] {Balance,Rate, Spread,Maturity,Resetting,FixedOrARM,PIOrIO,LiborCurve, "MortgageLoan_create" },
                    (objectType, parameters) =>
                    {
                        IMortgageLoan loan = construct_loan(Balance, Rate, Spread, Maturity, Resetting, FixedOrARM, PIOrIO, LiborCurve);
                        if (loan == null)
                            return ExcelError.ExcelErrorNull;
                        else
                            return loan;
                    });
            }
        }

        private static IMortgageLoan construct_loan(double Balance,double Rate,double Spread,int Maturity,int Resetting, string FixedOrARM, string PIOrIO, string Libor_Name)
        {
            LiborRates LiborRate;
            GlobalCache.TryGetObject<LiborRates>(Libor_Name, out LiborRate);

            IMortgageLoan loan = null;

            try
            {
                // Build the loan
                IRepayment rmPI = RepaymentFactory.GetRep(PIOrIO);
                loan            = MortgageLoanFactory.GetLoan(FixedOrARM, Balance, Maturity, Rate, Resetting, Spread, LiborRate, rmPI);
            }
            catch (Exception ex)
            {
                lock (m_sync)
                {
                    LogDisplay.WriteLine("Error: " + ex.ToString());
                }
            }
            return loan;
        }

        [ExcelFunction(Description = "Compute Mortgage Loan Cash Flows")]
        public static object[,] MortgageLoan_CashFlows([ExcelArgument(Description = @"MortgageLoan Object")] string loan_)
        {
            IMortgageLoan loanout;
            GlobalCache.TryGetObject<IMortgageLoan>(loan_, out loanout);

            loanout.CashFlows();

            double[] BegBal       = loanout.Write(loanout.ReturnBegBalance());
            double[] Interest     = loanout.Write(loanout.ReturnInterest());
            double[] Principal    = loanout.Write(loanout.ReturnPrincipal());
            double[] Collections  = loanout.Write(loanout.ReturnCashCollections());
            double[] EndBal       = loanout.Write(loanout.ReturnEndBalance());

            object[,] a = new object[MBSExcelDNA.Global.GlobalVar.GlobalMaxMortgageLoanMaturity, 5];
            for (int i = 0; i < (int)MBSExcelDNA.Global.GlobalVar.GlobalMaxMortgageLoanMaturity; i++)
            {
                a[i, 0] = BegBal[i];
                a[i, 1] = Interest[i];
                a[i, 2] = Principal[i];
                a[i, 3] = Collections[i];
                a[i, 4] = EndBal[i];
            }
            return a;
        }
    }
}

