// DumpARX.h: interface for the CDumpARX class.
//
//////////////////////////////////////////////////////////////////////

#if !defined(AFX_DUMPARX_H__705D831B_A9BD_4951_9D8D_6231EBA9F78C__INCLUDED_)
#define AFX_DUMPARX_H__705D831B_A9BD_4951_9D8D_6231EBA9F78C__INCLUDED_

#if _MSC_VER > 1000
#pragma once
#endif // _MSC_VER > 1000

#include "ARSConnection.h"
#include "FormList.h"
#include "RecordList.h"	// Added by ClassView

class AFX_EXT_CLASS CDumpARX  
{
public:
	int GetReturnCode();
	void SetStop(bool newbStop = true);
	BOOL DumpARXRecords(CARSConnection &arsConnect, CFormList &formList, 
				   CString strBackupDir, CString &strQualifier, CTime m_tModTime, UINT uiMaxNum = AR_NO_MAX_LIST_RETRIEVE);
	CDumpARX();
	virtual ~CDumpARX();
	enum return_code { rc_none = 0, /* uninitilized return code */
					   rc_stopped, /* backup job stopped */
					   rc_no_file /* output file couldn't be opened*/};
protected:
	bool CreateDirectorySeq(IN LPCTSTR lpszDirPath);
	int iReturnCode;
	CRecordList recordList;
	CString CleanupFileName(CString &strFilename);
	CString ReplaceFormChars(CString &strFormName);

};

#endif // !defined(AFX_DUMPARX_H__705D831B_A9BD_4951_9D8D_6231EBA9F78C__INCLUDED_)
