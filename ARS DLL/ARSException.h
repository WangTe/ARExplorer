// ARSException.h: interface for the CARSException class.
//
//////////////////////////////////////////////////////////////////////

#if !defined(AFX_ARSEXCEPTION_H__AAB30859_C330_4098_B52E_63386AB88A95__INCLUDED_)
#define AFX_ARSEXCEPTION_H__AAB30859_C330_4098_B52E_63386AB88A95__INCLUDED_

#if _MSC_VER > 1000
#pragma once
#endif // _MSC_VER > 1000

#define INDENT	"\t\t\t";

/////////////////////////////////////////////////////////////////////
// Global functions to throw ARException
void ThrowARSException(ARStatusList &StatusList, const char *pFunctionName);
void ThrowARSException(const char *cpText, const char *pFunctionName);
// global function to log errors to a file
//void Log(CString strOutputText);
void Log(const char *pOutputText);

class AFX_EXT_CLASS CARSException : public CException  
{

DECLARE_DYNAMIC(CARSException);

public:
	unsigned int uiType;
	int iErrorNum;
	CARSException(const char *pText);
	CARSException();
	CString strErrorText;
	CARSException(ARStatusList &pStatusList);
	virtual ~CARSException();
};

#endif // !defined(AFX_ARSEXCEPTION_H__AAB30859_C330_4098_B52E_63386AB88A95__INCLUDED_)
