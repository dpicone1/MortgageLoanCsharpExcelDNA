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

using MBSExcelDNA.Global;

namespace MBSExcelDNA.Loan
{
    public interface IMortgageLoan
    {
        void CashFlows();
        void PrintCashFlows();

        double[] Write(double[] v);

        double[] ReturnBegBalance();
        double[] ReturnEndBalance();
        double[] ReturnPrincipal() ;
        double[] ReturnInterest()  ;
        double[] ReturnCashCollections();
    }

    public abstract class MortgageLoan : IMortgageLoan
    {
        private IRepayment rep;          // repayment type of the loan {either P&I or IO}

        private LiborRates Libor_Curve;  // This is forward libor curve! In this project it is only used to load in the new rate when the ARM loan resets. 
                                         // In the future we can use it to calculate the loan PV based on a forward libor curve.  

        private double Loan_Rate;        // Loan rate 
        private double Original_Loan_Rate;// Original Loan rate
        private double Balance;          // Initial Balance 
        private int    Loan_Maturity;    // maturity 
        private double pmt;              // the monthly repayment

        private double[] Beg_Balance;
        private double[] End_Balance;
        private double[] Principal_Payment;
        private double[] Cash_Collections;
        private double[] Interest_Payment;

        public LiborRates LiborCurve
        {
            get { return this.Libor_Curve; }
        }

        public IRepayment Repayment
        {
            get { return this.rep; }            
        }

        public int LoanMaturity
        {
            get { return this.Loan_Maturity; }
        }

        public double PMT
        {
            get { return this.pmt; }
            set { this.pmt = value; }
        }

        public double LoanRate {
            get { return this.Loan_Rate; }
            set { this.Loan_Rate = value;}
        }

        public double OriginalLoanRate
        {
            get { return this.Original_Loan_Rate; }
        }
        //public double this[int i]
        //{
        //    get { return Beg_Balance[i]; }
        //    set { Beg_Balance[i] = value; }
        //}
        public double[] BegBalance       { get { return Beg_Balance; } }
        public double[] EndBalance       { get { return End_Balance; } }
        public double[] PrincipalPayment { get { return Principal_Payment; } }        
        public double[] InterestPayment  { get { return Interest_Payment; } }
        public double[] CashCollections  { get { return Cash_Collections; } }

        public double[] ReturnBegBalance()     { return BegBalance; }
        public double[] ReturnEndBalance()     { return EndBalance; }
        public double[] ReturnPrincipal()      { return PrincipalPayment; }
        public double[] ReturnInterest()       { return InterestPayment; }
        public double[] ReturnCashCollections(){ return CashCollections; }


        public MortgageLoan(double Balance_, int LoanMaturity_, double loanRate_, LiborRates LiborRate_Array_, IRepayment rep_)
        {
            this.rep          = rep_;
            this.Libor_Curve  = LiborRate_Array_;

            this.Balance      = Balance_;
            this.Loan_Maturity = LoanMaturity_;
            this.Loan_Rate     = loanRate_;
            this.Original_Loan_Rate = loanRate_;

            Beg_Balance       = new double[GlobalVar.GlobalMaxMortgageLoanMaturity];
            End_Balance       = new double[GlobalVar.GlobalMaxMortgageLoanMaturity];
            Principal_Payment = new double[GlobalVar.GlobalMaxMortgageLoanMaturity];
            Interest_Payment  = new double[GlobalVar.GlobalMaxMortgageLoanMaturity];
            Cash_Collections  = new double[GlobalVar.GlobalMaxMortgageLoanMaturity];
        }

        public abstract void Amortization(int RemainingPeriods);

        public double CalculatePmt(int CurrentPeriod, int RemainingPeriods)
        {
            return this.Repayment.pmt(this.BegBalance[CurrentPeriod], this.LoanRate, RemainingPeriods);
        }

        public void CashFlows()
        {
            for (int i = 0; i < this.Loan_Maturity; i++)
            {
                if (i == 0)
                    Beg_Balance[i] = Balance;
                else
                    Beg_Balance[i] = End_Balance[i - 1];

                Amortization(i);
            }
        }

        public void PrintCashFlows()
        {
            Console.WriteLine("{0, -8} {1, -5} {2, -5} {3, -7} {4, -5}\n", "BegBal", "Int", "Prin", "Coll", "EndBal");
            for (int i = 0; i < this.Loan_Maturity; i++)
            {
                Console.WriteLine("{0:f3} {1:f3} {2:f3} {3:f4} {4:f5}",
                    this.Beg_Balance[i],
                    this.Interest_Payment[i],
                    this.Principal_Payment[i],
                    this.Cash_Collections[i],
                    this.End_Balance[i]);
            }
        }

        public double[] Write(double[] vettore)
        {
            int len = vettore.Length;
            double[] v1 = new double[len];

            for (int i = 0; i < len; i++)
            { v1[i] = vettore[i]; }
            return v1;
        }
    }

    public class MortgageLoanARM : MortgageLoan
    {
        private int    NextRepricing;   // loan next repricing date 
        private double spread;          // loan spread over Libor rate

        public MortgageLoanARM(double Balance_, int LoanMaturity_, double loanRate_, int NextRepricing_, double spread_, LiborRates LiborRate_Array_, IRepayment rep_) :
            base(Balance_, LoanMaturity_, loanRate_, LiborRate_Array_, rep_)
        {
            this.NextRepricing = NextRepricing_;
            this.spread        = spread_;
        }

        public override void Amortization(int CurrentPeriod)
        {
            if(CurrentPeriod >= this.NextRepricing)
            {
                this.LoanRate = this.LiborCurve.LiborArray[CurrentPeriod] + this.spread;
            }

            // PMT calculation can be made more efficient by updating the New PMT only when the rate changes
            this.PMT = this.CalculatePmt(CurrentPeriod, this.LoanMaturity - CurrentPeriod);            

            this.InterestPayment[CurrentPeriod]  = this.BegBalance[CurrentPeriod] * this.LoanRate / 12;
            this.PrincipalPayment[CurrentPeriod] = System.Math.Max(0, this.PMT - this.InterestPayment[CurrentPeriod]);
            this.EndBalance[CurrentPeriod]       = System.Math.Max(0, this.BegBalance[CurrentPeriod] - this.PrincipalPayment[CurrentPeriod]);
            this.CashCollections[CurrentPeriod]  = this.InterestPayment[CurrentPeriod] + this.PrincipalPayment[CurrentPeriod];

            //After the last payment, reset the loan rate to the original loan rate
            if (CurrentPeriod == this.LoanMaturity - 1) this.LoanRate = this.OriginalLoanRate;
        }

    }
    public  class MortgageLoanFixed : MortgageLoan
    {
        public MortgageLoanFixed(double Balance_, int LoanMaturity_, double loanRate_, int NextRepricing_, double spread_, 
            LiborRates LiborRate_Array_ , IRepayment rep_) :
            base(Balance_, LoanMaturity_, loanRate_, LiborRate_Array_, rep_){ }

        public override void Amortization(int CurrentPeriod)
        {
            // PMT is calculated only once at the beginning as the loan rate never changes.            
            if (CurrentPeriod == 0)                     this.PMT = this.CalculatePmt(CurrentPeriod, this.LoanMaturity);
            
            // However, on the last period we recalculate the PMT as the fixed rate loan might be on an IO repayment scheme
            // This can be done more efficiently!! [to be amended next version]
            if (CurrentPeriod == this.LoanMaturity - 1) this.PMT = this.CalculatePmt(CurrentPeriod, this.LoanMaturity - CurrentPeriod);                


            this.InterestPayment[CurrentPeriod]  = this.BegBalance[CurrentPeriod] * this.LoanRate / 12;
            this.PrincipalPayment[CurrentPeriod] = System.Math.Max(0, this.PMT - this.InterestPayment[CurrentPeriod]);
            this.EndBalance[CurrentPeriod]       = System.Math.Max(0, this.BegBalance[CurrentPeriod] - this.PrincipalPayment[CurrentPeriod]);
            this.CashCollections[CurrentPeriod]  = this.InterestPayment[CurrentPeriod] + this.PrincipalPayment[CurrentPeriod];
        }
    }
};
