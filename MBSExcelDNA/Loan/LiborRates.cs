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
using System.IO;
using System.Reflection;

namespace MBSExcelDNA.Loan
{
    public class LiborRates
    {
        private double[] Libor_Array;

        public double[] LiborArray { get { return Libor_Array; } }
        public double this[int i]
        {
            get { return Libor_Array[i]; }
        }

        private StreamReader SetUpFileLocation()
        {
            var currentDirectory = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location);
            string NewPath = Path.GetFullPath(Path.Combine(currentDirectory, @"..\..\"));
            StreamReader reader = new StreamReader(File.OpenRead(NewPath + "LiborRates.csv"));
            return reader;
        }
        public LiborRates() { }

        public LiborRates(double[] Libor_Curve){
            int size         = Libor_Curve.Length;
            this.Libor_Array = new double[size];

            for (int i = 0; i < size; i++) Libor_Array[i] = Libor_Curve[i];
        }

        public void LoadingRates()
        {
            StreamReader reader = SetUpFileLocation();
            var LiborColumn = new List<string>();
            while (!reader.EndOfStream)
            {
                var splits = reader.ReadLine().Split(',');
                LiborColumn.Add(splits[0]);
                // I can read more data columns
            }
            var LiborArray = LiborColumn.ToArray();

            int size = LiborColumn.Count;

            this.Libor_Array = new double[size];

            for (int i = 0; i < size; i++) Libor_Array[i] = double.Parse(LiborArray[i]);
        }
    }
}
