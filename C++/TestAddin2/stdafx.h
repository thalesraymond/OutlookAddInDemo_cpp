// stdafx.h : include file for standard system include files,
// or project specific include files that are used frequently,
// but are changed infrequently

#pragma once

#ifndef STRICT
#define STRICT
#endif

#include "targetver.h"

#define _ATL_APARTMENT_THREADED
#define _ATL_NO_AUTOMATIC_NAMESPACE

#define _ATL_CSTRING_EXPLICIT_CONSTRUCTORS	// some CString constructors will be explicit

#include "resource.h"
#include <atlbase.h>
#include <atlcom.h>
#include <atlctl.h>
#include <msxml2.h>
#include  <Mshtml.h>
#include 	<Winhttp.h>
#include <atlstr.h>
#include <atlsafe.h>
//#import "C:\Program Files\Common Files\DESIGNER\MSADDNDR.DLL" raw_interfaces_only, raw_native_types, no_namespace, named_guids, auto_search ,auto_rename

// ###00006
// >> Adding a Ribbon and Ribbon Button
//----------------------------------------------

// See: http://blogs.msmvps.com/carlosq/2013/04/06/msaddndr-dll-file-not-installed-by-office-2013/

// This demo was created and last modified in 2011.  Since then,
// Microsoft has stopped including MSADDNDR.DLL with Office.  Bitrot.
// To make this demo work, MSADDNDR.DLL would need to be included by
// the addin installer.

// the fact that the direct import (above) is replaced by calling the
// library ID in the registry ( below ) is for the moment mystifying.

// See also: KB 2792179 on including the library manually.
// https://support.microsoft.com/en-us/kb/2792179
//-----------------------------------------------------------------------

//#import "C:\Program Files\Microsoft Office\Office14\MSOUTL.OLB" raw_interfaces_only, raw_native_types,  named_guids, auto_search, auto_rename,rename_namespace("Outlook")
//#import "C:\Program Files\Common Files\Microsoft Shared\OFFICE14\MSO.DLL" raw_interfaces_only, named_guids, auto_search, auto_rename,rename_namespace("Office")

// Office type library (i.e. mso.dll)
#import "libid:2DF8D04C-5BFA-101B-BDE5-00AA0044DE52"\
	auto_rename auto_search raw_interfaces_only rename_namespace("Office")
// Outlook type library (i.e. msoutl.olb)
#import "libid:00062FFF-0000-0000-C000-000000000046"\
	auto_rename auto_search raw_interfaces_only rename_namespace("Outlook")

// Forms type library (i.e. fm20.dll)
#import "libid:0D452EE1-E08F-101A-852E-02608C4D0BB4"\
	auto_rename auto_search raw_interfaces_only  rename_namespace("Forms")
using namespace Outlook;
using namespace Office;
using namespace Forms;
//add in type library 
#import "libid:AC0714F2-3D04-11D1-AE7D-00A0C90F26F4"\
	auto_rename auto_search raw_interfaces_only rename_namespace("AddinDesign")
using namespace AddinDesign;
#import "libid:00020905-0000-0000-C000-000000000046"\
	auto_rename auto_search raw_interfaces_only rename_namespace("Word")
////version("1.0") lcid("0")  raw_interfaces_only named_guids
//
//using namespace Word;

// ###00007
// >> Adding a Ribbon and Ribbon Button
//----------------------------------------------
// Note that the Ribbon extensibility interface may be
// "000C0396-0000-0000-C000-000000000046"
// See:
// https://msdn.microsoft.com/en-us/library/microsoft.office.core.iribbonextensibility.aspx?cs-save-lang=1&cs-lang=vb#code-snippet-1

// ###00009
// >> Adding a Ribbon and Ribbon Button
//----------------------------------------------
// The demo gets around the ribbon issue by importing the entire Office
// library: mso.dll
// The key to that is libid:2DF8D04C-5BFA-101B-BDE5-00AA0044DE52   
// just as used above.

//----------------------
// To find libid values:
//----------------------
// oleview.exe will provide you typelib (= libid) values.
//   You will, though, have to search for the library you are
//   importing.  oleview.exe is included in the Windows SDK, which you
//   may need to install.
//
// See: http://www.codeproject.com/Articles/3699/Importing-Type-Libraries

//---------------------------------------------------------------------------


using namespace ATL;

extern "C" const GUID ;