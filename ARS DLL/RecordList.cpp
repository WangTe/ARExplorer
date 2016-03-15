// RecordList.cpp: implementation of the CRecordList class.
//
//////////////////////////////////////////////////////////////////////

#include "stdafx.h"
#include "RecordList.h"
#include "ARSException.h"
#include "Record.h"

#ifdef _DEBUG
#undef THIS_FILE
static char THIS_FILE[]=__FILE__;
#define new DEBUG_NEW
#endif

//////////////////////////////////////////////////////////////////////
// Construction/Destruction
//////////////////////////////////////////////////////////////////////

CRecordList::CRecordList()
{
	bStop = false;
	::InitializeCriticalSection(&lock);
}

CRecordList::~CRecordList()
{
	::DeleteCriticalSection(&lock);
}

/////////////////////////////////////////////////////////////////
// Populates the list with all record entry-id's in the form.
// for joined forms, the entry-id's will have \n seperators
bool CRecordList::FillRecordList(CARSConnection &arsConnection, CString dumpForm, 
								 CString strQual, UINT uiMaxNum)
{
	RemoveAll(); // Initialize this class

	ARNameType formToGet; // IN: arsForm will go here
	ARQualifierStruct qualifier; // IN: strQual will go here
	ARQualifierStruct *pQualifier = NULL;
	AREntryListFieldList emptyGetListFields; // IN: this will be empty to not retrieve an additional fields
	AREntryListFieldValueList entryListFieldValueList; // OUT: this will hold the entry id list
	ARStatusList statusList; // OUT: the status list

	// Populate INPUT variables
	Form = dumpForm.Left(AR_MAX_NAME_SIZE);
	strcpy(formToGet, LPCSTR(Form)); // copy arsForm into formToGet

	if(!strQual.IsEmpty() ) {
		if(ARLoadARQualifierStruct(&arsConnection.LoginInfo,// Load the qualifier struct
										formToGet, 
										NULL, // VUI to use, default if NULL
										(char*)LPCSTR(strQual), 
										&qualifier, 
										&statusList) > AR_RETURN_OK)
		{
			FreeARQualifierStruct(&qualifier, FALSE);
			ThrowARSException(statusList, "CRecordList::GetRecords()");
		} // END IF

		pQualifier = &qualifier; // set pointer so get list call will use qualifier

		FreeARStatusList(&statusList, FALSE);// free it before we use again
	} // end if
	
	emptyGetListFields.numItems = 0; emptyGetListFields.fieldsList = NULL; // Init emptyGetListfields


	// return if stop command received
	::EnterCriticalSection(&lock);
	if(bStop) {
		FreeARQualifierStruct(&qualifier, FALSE);
		return false;
	}
	::LeaveCriticalSection(&lock);

	// Get the list of entry id's from the server
	if(ARGetListEntryWithFields(&arsConnection.LoginInfo,
								formToGet, // for name to get list of entry's from
								pQualifier, // qualification
								&emptyGetListFields, // empty list, don't get any fields (except entry id)
								NULL, // sort list, sorty by entry id
								0, // firstRetrieve, get the first entry
								uiMaxNum, // max number entries to get
								&entryListFieldValueList,
								NULL, // num matches
								&statusList) > AR_RETURN_OK)
	{	// handle the error
		FreeARQualifierStruct(&qualifier, FALSE);
		FreeAREntryListFieldValueList(&entryListFieldValueList, FALSE);
		ThrowARSException(statusList, "CRecordList::GetRecords()");
	}
	FreeARStatusList(&statusList, FALSE);

	// return if stop command received
	::EnterCriticalSection(&lock);
	if(bStop) {
		FreeARQualifierStruct(&qualifier, FALSE);
		FreeAREntryListFieldValueList(&entryListFieldValueList, FALSE);
		return false;
	}
	::LeaveCriticalSection(&lock);

	// Populate CRecordList with the Entry ID List
	CEntryId EntryId; // working Entry ID object
	AREntryListFieldValueStruct *pEntry = entryListFieldValueList.entryList;
	AREntryIdType *pEntryId; // working pointer
	// Loop once for each record returned
	for(UINT i=0; 
		i < entryListFieldValueList.numItems; 
		i++, pEntry++)
	{
		EntryId.RemoveAll(); // initialize
		CString tempString;

		// If only one entry id is in the list
		// save it go to next entry id structure
		if(pEntry->entryId.numItems == 0)
		{
			tempString = (char*)pEntry->entryId.entryIdList;
			EntryId.AddTail(tempString);
			AddTail(EntryId);
			continue;
		}else // add entry id from joined form
		{
			pEntryId = pEntry->entryId.entryIdList;
			tempString = (char*)pEntryId;
			EntryId.AddTail(tempString); // add the first entry id
			pEntryId++;

			// Loop once for each entry id in the entry
			// there will be multiple entry id's if the record
			// comes from a joined form
			for(UINT u=1;
				u<pEntry->entryId.numItems;
				u++, pEntry++)
				{
					tempString = (char*)pEntryId;
					EntryId.AddTail(tempString);
				}
			AddTail(EntryId);
		}// end else
	} // end for()

	FreeARQualifierStruct(&qualifier, FALSE);
	FreeAREntryListFieldValueList(&entryListFieldValueList, FALSE);
	return true;
}

////////////////////////////////////////////////////////////////////
// Returns a CRecord object of the record specified by POSITION
CRecord CRecordList::GetRecord(CARSConnection &arsConnect, POSITION pos,
							   CFieldList &FieldList)
{
	CRecord record; // record object to return
	ARNameType formToGet; // IN: form the get record for
	AREntryIdList entryId; // IN: the record's entry id that we're getting
	ARFieldValueList fieldList; // OUT: the record returned from ARGetEntry
	ARStatusList statusList; // OUT: list of errors
	CEntryId obEntryId; // working object

	// populate the entryId variable to save in record later
	record.EntryId = FillEntryId(&entryId, pos);

	// populate working variable formToGet
	strcpy(formToGet, (LPCSTR)Form);

	// Get the record from the server and store in fieldList
	if(ARGetEntry(&arsConnect.LoginInfo,
					formToGet,
					&entryId, // the entry id we just filled
					(ARIILP)FieldList, // IN: Ordered list of fields to get
					&fieldList,
					&statusList) > AR_RETURN_OK)
	{
		record.EntryId.RemoveAll();
		record.RemoveAll();
		ThrowARSException(statusList, "CRecordList::GetRecord()");
		return record;
		
	}
	FreeARStatusList(&statusList, FALSE);

// check out fieldList to make sure it was populated ok

	// Copy fieldList into record, no need to copy Entry Id, it was copied above
	record = fieldList;

	// Free heap memory
	FreeARFieldValueList(&fieldList, FALSE);
	FreeAREntryIdList(&entryId, FALSE);

	// return record
	return record;
}

CEntryId CRecordList::FillEntryId(AREntryIdList *pEntryIdList, POSITION pos)
{
	CEntryId SourceEntryId;
	SourceEntryId = GetAt(pos); // get the entry id and store locally.
	pEntryIdList->numItems = SourceEntryId.GetCount(); // store how many id's will make up this entry id list
	unsigned int size = sizeof(AREntryIdType) * SourceEntryId.GetCount();
	pEntryIdList->entryIdList = (AREntryIdType*)malloc(size); // allocated memory for entryId

	if(pEntryIdList->entryIdList == NULL)	// ensure memory was allocated
	{
		pEntryIdList->numItems = 0;
		AfxThrowMemoryException();
		SourceEntryId.RemoveAll();
		return SourceEntryId; // return empty entry id because of error
	}

	POSITION pos2 = SourceEntryId.GetHeadPosition();
	if(!pos2)	// EntryId was empty
	{
		ThrowARSException(ENTRY_ID_EMPTY, "CRecordList::FillEntryId()");
		SourceEntryId.RemoveAll();
		return SourceEntryId; // return empty entry id because of error
	}

	// now fill pEntryIdList with the entry id's in the current CEntryId object
	AREntryIdType *pEntryId = pEntryIdList->entryIdList;
	for(unsigned int i=0;
		i<pEntryIdList->numItems; 
		i++, pEntryId++)
	{
		strcpy((char*)pEntryId, LPCSTR(SourceEntryId.GetNext(pos2)) );
	}
	
	// return the CEntryId object in position "pos"
	return SourceEntryId;
}

/////////////////////////////////////////////////////////////////////////////////////////////
// Description: Dumps all AR records in the list.  Stops if bStop == true
// Returns:		true - if all records were dumped successfully
//				false - if dump terminated prematurely
bool CRecordList::DumpARX(CARSConnection &arsConnect, CStdioFile &File,
					 CString &strAttachDir, CTime m_tModTime)
{
	CString strBuffer; // working string buffer for faster file output
	CRecord record; // working record object
	unsigned int uiAttachNum = 0; // working count of attachments
	CFieldList FieldList(Form); // working ordered list of fields

	// Ensure the backup directory exists, if not create it
	CString strBackupPath(File.GetFilePath());	// get the full path to output file
	strBackupPath.Replace((LPCTSTR)File.GetFileName(), ""); // take out the file name
	
	// Create the attachment directory for this form
	strBackupPath += strAttachDir; 
	CreateDirectory((LPCTSTR)strBackupPath, NULL);

	// First get all data field id's in FieldList
	if(!FieldList.FillFieldList(arsConnect, AR_FIELD_TYPE_DATA))
		return false;

// ADD only output the header, if we're doing complete backup
	if(m_tModTime.GetTime() == 0) {
		// Output the ARX Header Information
		DumpARXHeader(arsConnect, File, FieldList);
	}

	// Output each record and attachments
	POSITION pos = GetHeadPosition();
	for(int i=0; i<GetCount(); i++, GetNext(pos))
	{
		// return if stop command received
		bool bTemp;
		::EnterCriticalSection(&lock);
		bTemp = bStop;
		::LeaveCriticalSection(&lock);
		if(bTemp) {
			File.Flush();
			return false;
		}

		// dump the current record and flush buffer
		record = GetRecord(arsConnect, pos, FieldList);
		record.DumpARX(arsConnect, Form, File, strAttachDir, &uiAttachNum, strBuffer);
		File.Write((LPCSTR)strBuffer, strBuffer.GetLength());
		File.Flush(); // could put this at end to speed up process
	}

	// return success
	return true;
}


CRecordList::DumpARXHeader(CARSConnection &arsConnect, 
						   CStdioFile &File, CFieldList &FieldList)
{
	CString strId;
	CString strBuffer;
	CString strFieldBuffer("FIELDS"), 
			strIdBuffer("FLD-ID"), 
			strTypeBuffer("DTYPES"); // working buffers
	POSITION pos; // working position
	CField Field; // working field object

	// Output SCHEMA "Form Name"<carriage return> to ARX File
	strBuffer = "SCHEMA \"" + Form + "\"\n";
	File.Write((LPCSTR)strBuffer, strBuffer.GetLength());
	File.Flush();
	strBuffer.Empty();

	// Loop through the field list outputing names, id's, & types 
	// into respective buffers
	pos = FieldList.GetHeadPosition(); // init pos
	for(int i=0; i<FieldList.GetCount(); i++)
	{
		Field = FieldList.GetNextField(arsConnect, pos);
		// Store field name
		strFieldBuffer += " \"" + Field.Name + "\"";

		// Store field id's 
		strId.Format("%d", Field.Id);
		strIdBuffer += " " + strId;

		// Store field data types
		strTypeBuffer += " " + Field.TypeText;
	}

	// now just output and flush the buffers in order
	strFieldBuffer += NEW_LINE;
	File.Write((LPCSTR)strFieldBuffer, strFieldBuffer.GetLength());
	File.Flush();

	strIdBuffer += NEW_LINE;
	File.Write((LPCSTR)strIdBuffer, strIdBuffer.GetLength());
	File.Flush();

	strTypeBuffer += NEW_LINE;
	File.Write((LPCSTR)strTypeBuffer, strTypeBuffer.GetLength());
	File.Flush();
}

////////////////////////////////////////////////////////////////////////
// Description: Set to true to stop a currently running job.
//				Set to false prior to running a job
void CRecordList::SetStop(bool newbStop)
{
	::EnterCriticalSection(&lock);

		bStop = newbStop;

	::LeaveCriticalSection(&lock);
}
