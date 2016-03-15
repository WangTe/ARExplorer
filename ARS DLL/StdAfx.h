// stdafx.h : include file for standard system include files,
//  or project specific include files that are used frequently, but
//      are changed infrequently
//

#if !defined(AFX_STDAFX_H__D4BDBA0F_E0C1_414E_B9FE_091F1B9FFFC8__INCLUDED_)
#define AFX_STDAFX_H__D4BDBA0F_E0C1_414E_B9FE_091F1B9FFFC8__INCLUDED_

#if _MSC_VER > 1000
#pragma once
#endif // _MSC_VER > 1000

#define VC_EXTRALEAN		// Exclude rarely-used stuff from Windows headers

#include <afxwin.h>         // MFC core and standard components
#include <afxext.h>         // MFC extensions

#ifndef _AFX_NO_OLE_SUPPORT
#include <afxole.h>         // MFC OLE classes
#include <afxodlgs.h>       // MFC OLE dialog classes
#include <afxdisp.h>        // MFC Automation classes
#endif // _AFX_NO_OLE_SUPPORT


#ifndef _AFX_NO_DB_SUPPORT
#include <afxdb.h>			// MFC ODBC database classes
#endif // _AFX_NO_DB_SUPPORT

#ifndef _AFX_NO_DAO_SUPPORT
#include <afxdao.h>			// MFC DAO database classes
#endif // _AFX_NO_DAO_SUPPORT

#include <afxdtctl.h>		// MFC support for Internet Explorer 4 Common Controls
#ifndef _AFX_NO_AFXCMN_SUPPORT
#include <afxcmn.h>			// MFC support for Windows Common Controls
#endif // _AFX_NO_AFXCMN_SUPPORT

#ifndef _C_LIST_TEMPLATE
#define _C_LIST_TEMPLATE
#include <afxtempl.h>
#endif // End _C_LIST_TEMPLATE

#ifndef __ARS_PRECOMPILED_HEADERS
#define __ARS_PRECOMPILED_HEADERS
#include <ar.h>
#include <arerrno.h>
#include <arextern.h>
#include <arstruct.h>
#include <arfree.h>

#endif // End __ARS_PRECOMPILED_HEADERS

//{{AFX_INSERT_LOCATION}}
// Microsoft Visual C++ will insert additional declarations immediately before the previous line.

#endif // !defined(AFX_STDAFX_H__D4BDBA0F_E0C1_414E_B9FE_091F1B9FFFC8__INCLUDED_)
