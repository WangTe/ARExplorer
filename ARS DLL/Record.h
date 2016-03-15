// Record.h: interface for the CRecord class.
//
//////////////////////////////////////////////////////////////////////

#if !defined(AFX_RECORD_H__F5F69828_D86E_4375_BE26_F5773B997DBA__INCLUDED_)
#define AFX_RECORD_H__F5F69828_D86E_4375_BE26_F5773B997DBA__INCLUDED_

#if _MSC_VER > 1000
#pragma once
#endif // _MSC_VER > 1000

#include "ARSConnection.h"


#define ARX_VALUE_SEPERATOR 4 // seperator character for Currency fields
#define ARX_HEADER_DATA "DATA" // header for a data row in ARX file


////////////////////////////////////////////////////////
// CRecord error messages
#define UNKNOWN_DATA_TYPE	"An unknown data type was encountered while trying to save the record."
#define CREC_ENTRY_ID_EMPTY	"The entry id list was empty."


////////////////////////////////////////////////////////
// List of valid data types
//--------------------------
// 2: Integer (AR_DATA_TYPE_INTEGER).
// 3: Real (AR_DATA_TYPE_REAL).
// 4: Character (AR_DATA_TYPE_CHAR).
// 5: Diary (AR_DATA_TYPE_DIARY).
// 6: Selection (AR_DATA_TYPE_ENUM).
// 7: Date/time (AR_DATA_TYPE_TIME).
// 10: Fixed-point decimal (AR_DATA_TYPE_DECIMAL). Values must be
//	   specified in C locale, for example 1234.56
// 11: Attachment (AR_DATA_TYPE_ATTACH).
// 12: Currency (AR_DATA_TYPE_CURRENCY).

class  AFX_EXT_CLASS CEntryId : public CStringList  
{
public:
	CEntryId(const CEntryId &Copy);
	CEntryId();
	virtual ~CEntryId();
	CEntryId CEntryId::operator +=(const CString &newEntryId);
//	CEntryId operator =(CEntryId &newEntryId);
	CEntryId operator =(const CEntryId &newEntryId);
};

////////////////////////////////////////////////////////
// Object to hold a single fields value and data type
// in the form of a String(value) and UINT(data type)
class  AFX_EXT_CLASS CFieldValuePair  
{
public:
	DumpAttachment(CARSConnection &arsConnect, CString Form, CEntryId EntryId, 
				   ARInternalId arsFieldId, CString strFilePath);
	DumpARX(CARSConnection &arsConnect, CString Form, CStdioFile &File,
			CString &strAttachDir, unsigned int *p_uiAttachNum,
			CEntryId &EntryId, CString &strBuffer);
	CString StoreCurrency(UINT uiNewFieldId, ARCurrencyStruct *currencyStruct);
	CString StoreAttachment(ARAttachStruct *attachment);
	ARInternalId uiFieldId;
	CFieldValuePair operator =(CFieldValuePair &newPair);
	CFieldValuePair(CString newValue, unsigned int uiNewType, ARInternalId fieldId);
	unsigned int uiType;
	CString Value;
	CFieldValuePair();
	virtual ~CFieldValuePair();

protected:
	int FillARLocStruct(CString &strFilePath, ARLocStruct &arsLocStruct);
	int FillEntryId(CEntryId &EntryId, AREntryIdList *p_arsEntryId);
};

///////////////////////////////////////////////////////
// Object to hold values about a single record
class AFX_EXT_CLASS CRecord  : public CList<CFieldValuePair, CFieldValuePair&>
{
public:
	CEntryId EntryId;
	DumpARX(CARSConnection &arsConnect, CString Form, CStdioFile &File,
			CString &strAttachDir, unsigned int *p_uiAttachNum, CString &strBuffer);
	CRecord(const CRecord &Copy);
	CRecord();
	virtual ~CRecord();
	CRecord operator =(CRecord &newRecord);
	CRecord operator =(ARFieldValueList &fieldList);
};




//class CAttachment  
//{
//public:
//
//};

#endif // !defined(AFX_RECORD_H__F5F69828_D86E_4375_BE26_F5773B997DBA__INCLUDED_)
