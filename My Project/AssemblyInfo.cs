﻿using System;
using System.Reflection;
using System.Runtime.InteropServices;

// General Information about an assembly is controlled through the following 
// set of attributes. Change these attribute values to modify the information
// associated with an assembly.

// Review the values of the assembly attributes

[assembly: AssemblyTitle("VSTO Addins")]
[assembly: AssemblyDescription("")]
[assembly: AssemblyCompany("")]
[assembly: AssemblyProduct("VSTO Addins")]
[assembly: AssemblyCopyright("Copyright ©  2023")]
[assembly: AssemblyTrademark("")]

// Setting ComVisible to false makes the types in this assembly not visible 
// to COM components.  If you need to access a type in this assembly from 
// COM, set the ComVisible attribute to true on that type.
[assembly: ComVisible(false)]

// The following GUID is for the ID of the typelib if this project is exposed to COM
[assembly: Guid("c89a61dc-bde3-4f2a-969e-bea84dbbdb97")]

// Version information for an assembly consists of the following four values:
// 
// Major Version
// Minor Version 
// Build Number
// Revision
// 
// You can specify all the values or you can default the Build and Revision Numbers 
// by using the '*' as shown below:
// <Assembly: AssemblyVersion("1.0.*")> 

[assembly: AssemblyVersion("1.0.0.0")]
[assembly: AssemblyFileVersion("1.0.0.0")]

namespace VSTO_Addins
{

    internal static class DesignTimeConstants
    {
        public const string RibbonTypeSerializer = "Microsoft.VisualStudio.Tools.Office.Ribbon.Serialization.RibbonTypeCodeDomSerializer, Microsoft.VisualStudio.Tools.Office.Designer, Version=10.0.0.0, Culture=neutral, PublicKeyToken=b03f5f7f11d50a3a";
        public const string RibbonBaseTypeSerializer = "System.ComponentModel.Design.Serialization.TypeCodeDomSerializer, System.Design";
        public const string RibbonDesigner = "Microsoft.VisualStudio.Tools.Office.Ribbon.Design.RibbonDesigner, Microsoft.VisualStudio.Tools.Office.Designer, Version=10.0.0.0, Culture=neutral, PublicKeyToken=b03f5f7f11d50a3a";
    }
}