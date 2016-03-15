// ARSConnection.h: interface for the CARSConnection class.
//
//////////////////////////////////////////////////////////////////////

#if !defined(AFX_ARSCONNECTION_H__2FB01F6C_0CE1_4ECF_B5DF_96A132F04699__INCLUDED_)
#define AFX_ARSCONNECTION_H__2FB01F6C_0CE1_4ECF_B5DF_96A132F04699__INCLUDED_

#if _MSC_VER > 1000
#pragma once
#endif // _MSC_VER > 1000

//////////////////////////////////////////////////////////////////////
// Custom Error Messages for CARSConnection Class
#define ERR_NO_USER				"There is no user specified."
#define ERR_NO_SERVER			"There is no server specified."

#include <ar.h>

/////////////////////////////////////////////////////////////////////
// Typedefs for conversion functions
typedef ARControlStruct* ARCSP;
typedef char* ARNT;
typedef ARNameType* ARNTP;

const char NEW_LINE='\n';

class AFX_EXT_CLASS CARSConnection : public CObject
{
	DECLARE_SERIAL(CARSConnection);
public:
	CString GetPassword();
	CString GetUser();
	BOOL IsLoggedIn();
	CARSConnection(CARSConnection &newConnect);
	CString GetServer();
	void Serialize( CArchive &archive );
	void Logout();
	void Login();
	void SetInfo(CString strUser, CString strPassword, CString strServer);
	CARSConnection(CString strUser, CString strPassword, CString strServer);
	CARSConnection();
	virtual ~CARSConnection();
	ARControlStruct LoginInfo;
	operator ARCSP();
	CARSConnection & operator =(CARSConnection &newConnect);
private:
	void ResetInfo();
};



#endif // !defined(AFX_ARSCONNECTION_H__2FB01F6C_0CE1_4ECF_B5DF_96A132F04699__INCLUDED_)
