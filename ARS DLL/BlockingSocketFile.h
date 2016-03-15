// BlockingSocketFile.h: interface for the CBlockingSocketFile class.
//
//////////////////////////////////////////////////////////////////////

#if !defined(AFX_BLOCKINGSOCKETFILE_H__61F83C6F_B0DD_4622_8004_36B0F9A70E12__INCLUDED_)
#define AFX_BLOCKINGSOCKETFILE_H__61F83C6F_B0DD_4622_8004_36B0F9A70E12__INCLUDED_

#if _MSC_VER > 1000
#pragma once
#endif // _MSC_VER > 1000

#include <Blocksock.h>

#define BLOCK_SOCK_TIMEOUT 15

class AFX_EXT_CLASS CBlockingSocketFile : public CFile  
{
	DECLARE_DYNAMIC(CBlockingSocketFile)
public:
//Constructors
	CBlockingSocketFile(CBlockingSocket *pNewSocket);
	virtual ~CBlockingSocketFile();

// Implementation
public:
	void RegisterLog( void (*newLog)(const char *) );
	CBlockingSocket * pSocket;

	virtual UINT Read(void * pBuffer, UINT nBytes);
	virtual void Write(const void* pBuffer, UINT nBytes);
	virtual void Close();

private:
	void (*Log)(const char *); // pointer to logging function
// Unsupported APIs
	virtual BOOL Open(LPCTSTR lpszFileName, UINT nOpenFlags, CFileException* pError = NULL);
	virtual CFile* Duplicate() const;
	virtual DWORD GetPosition() const;
	virtual LONG Seek(LONG lOff, UINT nFrom);
	virtual void SetLength(DWORD dwNewLen);
	virtual DWORD GetLength() const;
	virtual void LockRange(DWORD dwPos, DWORD dwCount);
	virtual void UnlockRange(DWORD dwPos, DWORD dwCount);
	virtual void Flush();
	virtual void Abort();
};

#endif // !defined(AFX_BLOCKINGSOCKETFILE_H__61F83C6F_B0DD_4622_8004_36B0F9A70E12__INCLUDED_)

/*
/////////////////////////////////////////////////////////////////////////////
// CSocketFile

class CSocketFile : public CFile
{
	DECLARE_DYNAMIC(CSocketFile)
public:
//Constructors
	CSocketFile(CSocket* pSocket, BOOL bArchiveCompatible = TRUE);

// Implementation
public:
	CSocket* m_pSocket;
	BOOL m_bArchiveCompatible;

	virtual ~CSocketFile();

#ifdef _DEBUG
	virtual void AssertValid() const;
	virtual void Dump(CDumpContext& dc) const;
#endif
	virtual UINT Read(void* lpBuf, UINT nCount);
	virtual void Write(const void* lpBuf, UINT nCount);
	virtual void Close();

// Unsupported APIs
	virtual BOOL Open(LPCTSTR lpszFileName, UINT nOpenFlags, CFileException* pError = NULL);
	virtual CFile* Duplicate() const;
	virtual DWORD GetPosition() const;
	virtual LONG Seek(LONG lOff, UINT nFrom);
	virtual void SetLength(DWORD dwNewLen);
	virtual DWORD GetLength() const;
	virtual void LockRange(DWORD dwPos, DWORD dwCount);
	virtual void UnlockRange(DWORD dwPos, DWORD dwCount);
	virtual void Flush();
	virtual void Abort();
};

*/