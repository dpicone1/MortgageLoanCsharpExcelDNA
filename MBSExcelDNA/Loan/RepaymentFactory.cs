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
    public class RepaymentFactory
    {
        private static Dictionary<string, Type> m_rep_types = new Dictionary<string, Type>();

        static RepaymentFactory()
        {
            Type base_type = typeof(IRepayment);

            Type[] types = Common.GetClassTypes(System.Reflection.Assembly.GetExecutingAssembly(), base_type.Namespace);

            foreach (Type t in types)
            {
                if (!t.IsAbstract && base_type.IsAssignableFrom(t))
                {
                    m_rep_types[t.FullName.ToLower()] = t;
                }
            }
        }

        public static Type GetRepType(string method)
        {
            string name = String.Format("MBSExcelDNA.Loan.Repayment{0}", method).ToLower();

            Type result;

            m_rep_types.TryGetValue(name, out result); //returs null if not found

            return result;
        }

        public static IRepayment GetRep(Type type, double x)
        {
            IRepayment rep = Activator.CreateInstance(type, x) as IRepayment;

            return rep;
        }

        public static IRepayment GetRep(Type type, params object[] arguments)
        {
            IRepayment rep = Activator.CreateInstance(type, arguments) as IRepayment;

            return rep;
        }

        public static IRepayment GetRep(string method, params object[] arguments)
        {
            IRepayment rep = null;

            Type rep_type = GetRepType(method);

            if (rep_type != null)
            {
                rep = Activator.CreateInstance(rep_type, arguments) as IRepayment;
            }

            return rep;
        }

        public static IRepayment GetRep(string method, double x)
        {
            IRepayment rep = null;

            Type rep_type = GetRepType(method);

            if (rep_type != null)
            {
                rep = Activator.CreateInstance(rep_type, x) as IRepayment;
            }

            return rep;
        }
    }

}
