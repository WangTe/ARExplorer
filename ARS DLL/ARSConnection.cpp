// ARSConnection.cpp: implementation of the CARSConnection class.
//
//////////////////////////////////////////////////////////////////////

#include "StdAfx.h"
#include "ARSConnection.h"
#include "ARSException.h"
#include "Globals.h"

#ifdef _DEBUG
#undef THIS_FILE
static char THIS_FILE[]=__FILE__;
#define new DEBUG_NEW
#endif


//////////////////////////////////////////////////////////////////////
// Static variable declerations.
//ARControlStruct		CARSConnection::LoginInfo;
IMPLEMENT_SERIAL(CARSConnection, CObject, 1);

//////////////////////////////////////////////////////////////////////
// Construction/Destruction
//////////////////////////////////////////////////////////////////////

CARSConnection::CARSConnection()
{
	ResetInfo();  // Init User, Password, & Server

	// Initialize static parameters
	LoginInfo.cacheId = 0;
	LoginInfo.sessionId = 0;
	LoginInfo.language[0] = NULL;
}

CARSConnection::~CARSConnection()
{
	Logout();
}

CARSConnection::CARSConnection(CString strUser, CString strPassword, CString strServer)
{
	CARSConnection::CARSConnection(); // Initialize variables

	SetInfo(strUser, strPassword, strServer);
}

void CARSConnection::SetInfo(CString strUser, CString strPassword, CString strServer)
{
	// If connected, first logout
	Logout();

	strUser = strUser.Left(AR_MAX_ACCESS_NAME_SIZE);
	strPassword = strPassword.Left(AR_MAX_ACCESS_NAME_SIZE);
	strServer = strServer.Left(AR_MAX_SERVER_SIZE);

	if(!strUser.IsEmpty()) // Ensure username was given
		strcpy(LoginInfo.user, LPCSTR(strUser));
	else
		ThrowARSException(ERR_NO_USER, "CARSConnection::SetInfo");

	if(strPassword.IsEmpty()) // Ensure password was given
		LoginInfo.password[0] = NULL;
	else
		strcpy(LoginInfo.password, LPCSTR(strPassword));

	if(!strServer.IsEmpty()) // Ensure server name was given
		strcpy(LoginInfo.server, LPCSTR(strServer));
	else
		ThrowARSException(ERR_NO_SERVER, "CARSConnection::SetInfo");
}

void CARSConnection::Login()
{
	ARStatusList StatusList;
	CString strError;

	// If connected, logout first.
	if( LoginInfo.sessionId != 0)
		Logout();

	// Check for Login Name
	if(LoginInfo.user[0] == NULL)
		ThrowARSException(ERR_NO_USER, "CARSConnection::Login()");

	// Check for server name
	if(LoginInfo.server[0] == NULL)
		ThrowARSException(ERR_NO_SERVER, "CARSConnection::Login()");

	// If we haven't initialized yet, then initialize the connection.
	if(!LoginInfo.cacheId)
	{
		/* Initialize Remedy session.  This call establishes the environment for	*/
		/* interaction with the AR System API.  It must be the first AR API call	*/
		/* made in every application.  It will populate the CacheId and the			*/
		/* SessionID.																*/
		if ( ARInitialization( &LoginInfo, &StatusList ) > AR_RETURN_OK )
		{
			// Store the last error and free the status list.
			ThrowARSException(StatusList, "CARSConnection::Login()");
		}
		FreeARStatusList(&StatusList, FALSE);

	}

	// Log into the ARS
	if(ARVerifyUser(&LoginInfo, NULL, NULL, NULL, &StatusList) > AR_RETURN_OK)
	{
		// Throw the error
		ThrowARSException(StatusList, "CARSConnection::Login()");
	}
	FreeARStatusList(&StatusList, FALSE);
}

// Should only be used in the constructor
void CARSConnection::ResetInfo()
{
	LoginInfo.user[0] = NULL;
	LoginInfo.password[0] = NULL;
	LoginInfo.server[0] = NULL;
	LoginInfo.sessionId = 0;
	LoginInfo.cacheId = 0;
}

void CARSConnection::Logout()
{
	ARStatusList StatusList;
	if(IsLoggedIn())
	{	// If logged in, log out from Remedy
		if(ARTermination(&LoginInfo, &StatusList) > AR_RETURN_OK)
		{
			ThrowARSException(StatusList, "CARSConnection::Logout()");
		}
		FreeARStatusList(&StatusList, FALSE);
	}
	ResetInfo();  // reset for next login
}

CARSConnection::operator ARCSP()
{
	return &LoginInfo;
}

//////////////////////////////////////////////////////////
// Description: serializes the object to a CArchive object
void CARSConnection::Serialize(CArchive &archive)
{
	// call base class serialize function first
	CObject::Serialize( archive );

	// now serialize this object
	if(archive.IsStoring()) {
		CString strUser(LoginInfo.user), strPassword(LoginInfo.password), 
				strServer(LoginInfo.server);
		archive << strUser;
		archive << strPassword;
		archive << strServer;
	}else {
		CString strUser, strPassword, strServer;
		archive >> strUser;
		archive >> strPassword;
		archive >> strServer;
		strcpy(LoginInfo.user, LPCTSTR(strUser));
		strcpy(LoginInfo.password, LPCTSTR(strPassword));
		strcpy(LoginInfo.server, LPCTSTR(strServer));
	}
	log.Log(CFileLogging::level_lower, "CARSConnection:Serialize(): Done serializing login info.", NULL);
}

CARSConnection & CARSConnection::operator =(CARSConnection &newConnect)
{
	LoginInfo = newConnect.LoginInfo;

	return *this;
}

CString CARSConnection::GetServer()
{
	CString server(LoginInfo.server);
	return server;
}

/////////////////////////////////////////////////////
// Returns: TRUE if logged in, FALSE if not logged in
//DEL BOOL CARSConnection::LoggedIn()
//DEL {
//DEL 	// If connected, logout first.
//DEL 	if( LoginInfo.sessionId != 0)
//DEL 		return TRUE;
//DEL 	else 
//DEL 		return FALSE;
//DEL }

CARSConnection::CARSConnection(CARSConnection &newConnect)
{
	operator=(newConnect);
}

BOOL CARSConnection::IsLoggedIn()
{
	if(LoginInfo.sessionId)
		return TRUE;
	else 
		return FALSE;
}

CString CARSConnection::GetUser()
{
	CString strUser = LoginInfo.user;
	return strUser;
}

CString CARSConnection::GetPassword()
{
	CString strPassword = LoginInfo.password;
	return strPassword;
}
