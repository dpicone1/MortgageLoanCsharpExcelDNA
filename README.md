# MortgageLoanCsharpExcelDNA

Building Mortgage Loan Cash flows with C# Excel DNA Add-ins

In this C# project we create a simple mortgage loan library to deal with the repayment of a loan. The library deals with two types of amortisation: Principal and Interest, and Interest Only. The loan interest rate can be fixed for life, or fixed for the first years, and then it resets to a variable rate plus a spread.

We use Excel DNA to expose our C# library to excel. Excel DNA provides an easy way to create C# functions which can then be used in Excel.

Our simple mortgage library can be used either via a VBA macro or as a standard excel function. Additionally, we also added an extra useful feature: the loan object can be stored in memory, and consequently run to produce cash flows. This is indeed very useful. Imagine you are dealing with a more complex task: pricing several interest rate option structures with the same interest rate simulation model. The only differentiation in each option is the payoff. You do not want to rerun the interest rate simulation each time you price a single option. This is very inefficient. Is it possible to call constructor once, store the simulation object somewhere and then use it to compute option price based upon its payoff? To the best of my knowledge, this solution was originally presented by Alex Chirokov in https://www.codeproject.com/Articles/1097174/Interpolation-in-Excel-using-Excel-DNA.

We reuse here his original implementation and allow the user to create a loan, store it, and then run the relative cashflows.

We have also attached the ExcelDNA manual by Govert van Drimmelen, https://excel-dna.net/, "Excel-DNA - Step-by-step C# add-in.doc"

Before changing the project, check it is working. Open the excel file "ExcelMortgages.xlsm" and the xll "MBSExcelDNA_ForDistribution.xll". 

You need ExcelDnaPack.exe only when releasing the project. You may never need to do so. Again read the doc file for more info.

In case you need more help with ExcelDNA, Mikael Katajamäki runs a very useful blog for everyone interested in using Excel DNA. https://mikejuniperhill.blogspot.com/2014/03/using-c-excel-addin-with-exceldna.html.
