using System.Reflection;
using System.Runtime.CompilerServices;

//
// General Information about an assembly is controlled through the following 
// set of attributes. Change these attribute values to modify the information
// associated with an assembly.
//
[assembly: AssemblyTitle("")]
#if VS2012
[assembly: AssemblyDescription("ModelMaker IDE integration Addin for VS 2012")]
#else
#if VS2010
[assembly: AssemblyDescription("ModelMaker IDE integration Addin for VS 2010")]
#else
#if VS2008
[assembly: AssemblyDescription("ModelMaker IDE integration Addin for VS 2008")]
#else
#if VS2005
[assembly: AssemblyDescription("ModelMaker IDE integration Addin for VS 2005")]
#else
[assembly: AssemblyDescription("ModelMaker IDE integration Addin for VS 2003")]
#endif
#endif
#endif
#endif
[assembly: AssemblyConfiguration("")]
[assembly: AssemblyCompany("ModelMaker Tools BV")]
[assembly: AssemblyProduct("ModelMaker 11")]
[assembly: AssemblyCopyright("(c) 2001-2012 ModelMaker Tools BV")]
[assembly: AssemblyTrademark("ModelMaker[R]")]
[assembly: AssemblyCulture("")]		

//
// Version information for an assembly consists of the following four values:
//
//      Major Version
//      Minor Version 
//      Revision
//      Build Number
//
// You can specify all the value or you can default the Revision and Build Numbers 
// by using the '*' as shown below:

[assembly: AssemblyVersion("1.0.*")]

//
// In order to sign your assembly you must specify a key to use. Refer to the 
// Microsoft .NET Framework documentation for more information on assembly signing.
//
// Use the attributes below to control which key is used for signing. 
//
// Notes: 
//   (*) If no key is specified - the assembly cannot be signed.
//   (*) KeyName refers to a key that has been installed in the Crypto Service
//       Provider (CSP) on your machine. 
//   (*) If the key file and a key name attributes are both specified, the 
//       following processing occurs:
//       (1) If the KeyName can be found in the CSP - that key is used.
//       (2) If the KeyName does not exist and the KeyFile does exist, the key 
//           in the file is installed into the CSP and used.
//   (*) Delay Signing is an advanced option - see the Microsoft .NET Framework
//       documentation for more information on this.
//
[assembly: AssemblyDelaySign(false)]
[assembly: AssemblyKeyFile("")]
[assembly: AssemblyKeyName("")]
