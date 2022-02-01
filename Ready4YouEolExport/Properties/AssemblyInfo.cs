using System.Reflection;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;

// General Information about an assembly is controlled through the following
// set of attributes. Change these attribute values to modify the information
// associated with an assembly.
[assembly: AssemblyTitle("Ready4YouEolExport")]
[assembly: AssemblyDescription("")]
[assembly: AssemblyConfiguration("")]
[assembly: AssemblyCompany("SolidOnline")]
[assembly: AssemblyProduct("Ready4YouEolExport")]
[assembly: AssemblyCopyright("Copyright SolidOnline ©  2020")]
[assembly: AssemblyTrademark("")]
[assembly: AssemblyCulture("")]

// Setting ComVisible to false makes the types in this assembly not visible
// to COM components.  If you need to access a type in this assembly from
// COM, set the ComVisible attribute to true on that type.
[assembly: ComVisible(false)]

// The following GUID is for the ID of the typelib if this project is exposed to COM
[assembly: Guid("fe5edebc-03c7-4b60-b003-fc122af2e3fc")]

// Version information for an assembly consists of the following four values:
//
//      Major Version
//      Minor Version
//      Build Number
//      Revision
//
// You can specify all the values or you can default the Build and Revision Numbers
// by using the '*' as shown below:
// [assembly: AssemblyVersion("1.0.*")]


//2020-11-10      1.0.0.0   Initial Version
//2020-11-12      1.0.0.1   Fallout logic change, Message changes 
//2020-11-19      1.0.1.0   CSV Format changes , EOL Output stored as file.
//2020-11-20      1.0.1.1   EOL Exceptions Catched.
//2020-11-24      1.0.2.0   NL Culture.
//2020-11-24      1.0.3.0   Additional Exceptions Handled.
//2020-11-27      1.0.3.0   Handling ',' in XML Generation
//2020-11-27      1.0.4.0   Handling ',' ONLY in XML Generation
//2020-11-30      1.0.5.0   Insert Empty row before the new costcenter starts 
//                          Ignore the csv records with aantal zero
//                          Special characters in names like é, ü, ä are not coming correctly in EOL
//2020-12-02      1.0.6.0   The rate from the UREN csv should be multiplied by the percentage 
//                          Refresh Token shound be available from a different path 
//2020-12-07      1.0.7.0   KM count should be rounded to INT directly after reading from the csv 
//                          Hourly rates:- take value from the cvs and round on 2 decimals
//                          Hourly rates:- multiply by the percentage and round again on 2 decimals
//                          CalculateUnitPrice :- Round on 2 decimals
//2020-12-16      1.0.8.0   when it comes to the fixed rate for temps of € 27.95 OR the urgent rate of € 29.95, then Articles 125 and 150 the surcharges of 125% and 150% should also be applied
//2020-12-21      1.0.9.0   Two other items:TELEFOON and CONSIGNATIE :
//                          Telefoon - price is in the csv, count always 1
//                          Consignatie - we need to have a factor in the tool
//2020-12-23      1.0.10.0  CONSIGNATIE : it should be count * rate from csv * factor 
//2021-06-17      1.0.11.0  Change in create access token method for EOL API 
//2021-11-10      1.0.12.0  Bugfix in create access token method for EOL API 
//2022-02-01      1.0.13.0  New Settings for Werksoort calculations using replacement werksoort. 
[assembly: AssemblyVersion("1.0.13.0")]
[assembly: AssemblyFileVersion("1.0.13.0")]