// RecordList.h: interface for the CRecordList class.
//
//////////////////////////////////////////////////////////////////////

#if !defined(AFX_RECORDLIST_H__38FFF596_2108_4717_8CE7_B47D61DD9354__INCLUDED_)
#define AFX_RECORDLIST_H__38FFF596_2108_4717_8CE7_B47D61DD9354__INCLUDED_

#if _MSC_VER > 1000
#pragma once
#endif // _MSC_VER > 1000

#include "ARSConnection.h"
#include "Form.h"
#include "Record.h"	// Added by ClassView

#define ENTRY_ID_SEPERATOR "\n"

//////////////////////////////////////////////////////////////////////
// CRecordList Custom Error Messages
#define ENTRY_ID_EMPTY	"The entry id list was empty."

class AFX_EXT_CLASS CRecordList : public CList<CEntryId, CEntryId&>
{
public:
	void SetStop(bool newbStop = true);
	bool DumpARX(CARSConnection &arsConnect, CStdioFile &File,
			CString &strAttachDir, CTime m_tModTime);
	CRecord GetRecord(CARSConnection &arsConnect, POSITION pos,
					  CFieldList &FieldList);
	bool FillRecordList(CARSConnection &arsConnection, CString dumpForm, 
						CString strQual, UINT uiMaxNum = AR_NO_MAX_LIST_RETRIEVE);
	CRecordList();
	virtual ~CRecordList();

protected:
	CRITICAL_SECTION lock;
	bool bStop;
	CEntryId FillEntryId(AREntryIdList *pEntryIdList, POSITION pos);
	CString Form;
	DumpARXHeader(CARSConnection &arsConnect, 
				  CStdioFile &File, CFieldList &FieldList);
};

#endif // !defined(AFX_RECORDLIST_H__38FFF596_2108_4717_8CE7_B47D61DD9354__INCLUDED_)
