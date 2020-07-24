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
    public class MortgageLoanFactory
    {
        private static Dictionary<string, Type> m_loan_types = new Dictionary<string, Type>();

        static MortgageLoanFactory()
        {
            Type base_type = typeof(IMortgageLoan);

            Type[] types = Common.GetClassTypes(System.Reflection.Assembly.GetExecutingAssembly(), base_type.Namespace);

            foreach (Type t in types)
            {
                if (!t.IsAbstract && base_type.IsAssignableFrom(t))
                {
                    m_loan_types[t.FullName.ToLower()] = t;
                }
            }
        }

        public static Type GetLoanType(string method)
        {
            string name = String.Format("MBSExcelDNA.Loan.MortgageLoan{0}", method).ToLower();

            Type result;

            m_loan_types.TryGetValue(name, out result); //returs null if not found

            return result;
        }

        public static IMortgageLoan GetLoan(Type type, double x)
        {
            IMortgageLoan loan = Activator.CreateInstance(type, x) as IMortgageLoan;

            return loan;
        }

        public static IMortgageLoan GetLoan(Type type, params object[] arguments)
        {
            IMortgageLoan loan = Activator.CreateInstance(type, arguments) as IMortgageLoan;

            return loan;
        }

        public static IMortgageLoan GetLoan(string method, params object[] arguments)
        {
            IMortgageLoan loan = null;

            Type loan_type = GetLoanType(method);

            if (loan_type != null)
            {
                loan = Activator.CreateInstance(loan_type, arguments) as IMortgageLoan;
            }

            return loan;
        }

        public static IMortgageLoan GetLoan(string method, double x)
        {
            IMortgageLoan loan = null;

            Type loan_type = GetLoanType(method);

            if (loan_type != null)
            {
                loan = Activator.CreateInstance(loan_type, x) as IMortgageLoan;
            }

            return loan;
        }
    }

}
