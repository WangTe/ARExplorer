// BlockingSocketFile.cpp: implementation of the CBlockingSocketFile class.
//
//////////////////////////////////////////////////////////////////////

#include "stdafx.h"
#include <Blocksock.h>
#include "BlockingSocketFile.h"

IMPLEMENT_DYNAMIC(CBlockingSocketFile, CFile);

//////////////////////////////////////////////////////////////////////
// Construction/Destruction
//////////////////////////////////////////////////////////////////////

CBlockingSocketFile::CBlockingSocketFile(CBlockingSocket *pNewSocket)
{
	if(pNewSocket == NULL) {
		AfxThrowUserException();
	}
	pSocket = pNewSocket;
	Log = NULL;
}

CBlockingSocketFile::~CBlockingSocketFile()
{
	// Added to try and fix bug in release build
//	CFile::~CFile();
}

UINT CBlockingSocketFile::Read(void *pBuffer, UINT nBytes)
{
	UINT uiCount = pSocket->Receive((char *)pBuffer, (const int)nBytes, BLOCK_SOCK_TIMEOUT);

	// Output to debugging log if it's registered
	if(Log) {
		CString text; text.Format("CBlockingSocketFile::Read(): read %d bytes", uiCount);
		Log((LPCTSTR)text);
	}
	return uiCount;
}

void CBlockingSocketFile::Write(const void *pBuffer, UINT nBytes)
{
	int i = 0;
	i = pSocket->Write((const char*)pBuffer, (const int)nBytes, BLOCK_SOCK_TIMEOUT);

	// Output to debugging log if it's registered
	if(Log) {
		CString text; text.Format("CBlockingSocketFile::Write(): wrote %d bytes", i);
		Log((LPCTSTR)text);
	}

} // End Write()

void CBlockingSocketFile::Close()
{
	pSocket->Close();
}

// Unsupported APIs
BOOL CBlockingSocketFile::Open(LPCTSTR lpszFileName, UINT nOpenFlags, CFileException* pError){ return false; }
CFile* CBlockingSocketFile::Duplicate() const {return NULL;}
DWORD CBlockingSocketFile::GetPosition() const {return 0;}
LONG CBlockingSocketFile::Seek(LONG lOff, UINT nFrom){return 0;}
void CBlockingSocketFile::SetLength(DWORD dwNewLen){}
DWORD CBlockingSocketFile::GetLength() const {return 0;}
void CBlockingSocketFile::LockRange(DWORD dwPos, DWORD dwCount){}
void CBlockingSocketFile::UnlockRange(DWORD dwPos, DWORD dwCount){}
void CBlockingSocketFile::Flush(){}
void CBlockingSocketFile::Abort(){}

void CBlockingSocketFile::RegisterLog(void (__cdecl *newLog)(const char *))
{
	Log = newLog;
}
