// Record.cpp: implementation of the CRecord class.
//
//////////////////////////////////////////////////////////////////////

#include "stdafx.h"
#include "Record.h"
#include "ARSException.h"

#ifdef _DEBUG
#undef THIS_FILE
static char THIS_FILE[]=__FILE__;
#define new DEBUG_NEW
#endif

//////////////////////////////////////////////////////////////////////
// Construction/Destruction
//////////////////////////////////////////////////////////////////////

CRecord::CRecord()
{

}

CRecord::~CRecord()
{

}

//////////////////////////////////////////////////////////////////////
// FieldValuePair Class
//////////////////////////////////////////////////////////////////////

//////////////////////////////////////////////////////////////////////
// Construction/Destruction
//////////////////////////////////////////////////////////////////////

CFieldValuePair::CFieldValuePair()
{
	uiType = 0;
	uiType = 0;
	Value.Empty();
}

CFieldValuePair::~CFieldValuePair()
{

}

CFieldValuePair::CFieldValuePair(CString newValue, unsigned int uiNewType, ARInternalId fieldId)
{
	uiType = uiNewType;
	Value = newValue;
	uiFieldId = fieldId;
}

CFieldValuePair CFieldValuePair::operator =(CFieldValuePair &newPair)
{
	uiFieldId = newPair.uiFieldId;
	uiType = newPair.uiType;
	Value = newPair.Value;
	return *this;
}

CRecord::CRecord(const CRecord &Copy)
{
	POSITION pos = Copy.GetHeadPosition();
	for(int i=0; i<Copy.GetCount(); i++)
	{
		AddTail(Copy.GetNext(pos));
	}
	EntryId = Copy.EntryId;
}

//CEntryId CEntryId::operator =(CEntryId &newEntryId)
//{
//	POSITION pos = newEntryId.GetHeadPosition();
//	for(int i=0; i<newEntryId.GetCount(); i++)
//	{
//		AddTail(newEntryId.GetNext(pos));
//	}
//	return *this;
//}

CEntryId CEntryId::operator =(const CEntryId &newEntryId)
{
#ifdef _DEBUG
	int numEntries = newEntryId.GetCount();
#endif

	// init CEntryId
	RemoveAll();

	POSITION pos = newEntryId.GetHeadPosition();
	for(int i=0; i<newEntryId.GetCount(); i++)
	{
		AddTail(newEntryId.GetNext(pos));
	}
	return *this;
}

CEntryId::CEntryId(const CEntryId &Copy)
{
	POSITION pos = Copy.GetHeadPosition();
	for(int i=0;
		i<Copy.GetCount();
		i++)
	{
		AddTail(Copy.GetNext(pos));
	}
}

CEntryId CEntryId::operator +=(const CString &newEntryId)
{
	AddTail(newEntryId);
	return *this;
}

CRecord CRecord::operator =(CRecord &newRecord)
{
	// init the record before copying.
	if(!IsEmpty())
		RemoveAll();

	// copy each field/value pair
	POSITION pos = newRecord.GetHeadPosition();
	for(int i=0; i<newRecord.GetCount(); i++)
	{
		AddTail(newRecord.GetNext(pos));
	}

	EntryId = newRecord.EntryId;

	return *this;
}

CRecord CRecord::operator =(ARFieldValueList &fieldList)
{
	ARFieldValueStruct *pFieldValuePair = fieldList.fieldValueList; // working pointer
	CFieldValuePair FieldValuePair; // working field value pair
	CString workingString; // working string to convert value

	// Loop once for each field/value pair in the list
	for(UINT i=0; i<fieldList.numItems; i++, pFieldValuePair++)
	{
		// decode the field/value pair
		FieldValuePair.uiFieldId = pFieldValuePair->fieldId; // save field id
		FieldValuePair.uiType = pFieldValuePair->value.dataType; // save data type
		switch(pFieldValuePair->value.dataType) // convert the value based on data type
		{
		case AR_DATA_TYPE_NULL:
			workingString.Empty();
			break;
		case AR_DATA_TYPE_INTEGER:
		case AR_DATA_TYPE_TIME:
		case AR_DATA_TYPE_DATE:
		case AR_DATA_TYPE_TIME_OF_DAY:
		case AR_DATA_TYPE_KEYWORD:
		case AR_DATA_TYPE_BITMASK:
			workingString.Format("%d", pFieldValuePair->value.u.intVal);
			break;
		case AR_DATA_TYPE_REAL:
			workingString.Format("%e", pFieldValuePair->value.u.realVal);
			break;
		case AR_DATA_TYPE_CHAR:
		case AR_DATA_TYPE_DIARY:
		case AR_DATA_TYPE_DECIMAL:
			workingString = pFieldValuePair->value.u.charVal;
			break;
		case AR_DATA_TYPE_ENUM:
			workingString.Format("%u", pFieldValuePair->value.u.enumVal);
			break;
		case AR_DATA_TYPE_ATTACH:
			workingString = FieldValuePair.StoreAttachment(pFieldValuePair->value.u.attachVal);
			break;
		case AR_DATA_TYPE_CURRENCY:
			workingString = FieldValuePair.StoreCurrency(pFieldValuePair->fieldId,
														 pFieldValuePair->value.u.currencyVal);
			break;
		default:
			workingString.Empty();
			ThrowARSException(UNKNOWN_DATA_TYPE, "CRecord::operator=()");
			break;
		} // end switch
		
		// decode attachement fields

		// decode currency fields
		FieldValuePair.Value = workingString; // save the decoded value

		AddTail(FieldValuePair); // Add the field/value pair
	} // end for
	
	// return the populated record
	return *this;
}

////////////////////////////////////////////////////////////////////////////////////////
// This function will store the file name only in Value as a text string.
CString CFieldValuePair::StoreAttachment(ARAttachStruct *attachment)
{
	CString	strWorking(attachment->name); // working string
	// decode attachment
	strWorking.MakeReverse();

	Value.Empty(); // init value to empty

	/* start from end, working backward stop at / or \ */
	for(int i=0; i<strWorking.GetLength(); i++)
	{
		TCHAR c;
		c = strWorking[i];
		if(c == '\\' || c == '/')	// if we reach the end of the actual filename
		{
			break;
		}else
			Value += c; // store the characters in backwards order
	} // end for

	Value.MakeReverse(); // put the filename in regular order

	return Value; // so you can do in line assignment

}

///////////////////////////////////////////////////////////////////////////
// Stores a currenty data type in Value
CString CFieldValuePair::StoreCurrency(UINT uiNewFieldId, ARCurrencyStruct *currencyStruct)
{
	CString strWork; // working string
	ARFuncCurrencyStruct *pFuncCurrency = currencyStruct->funcList.funcCurrencyList;
	uiFieldId = uiNewFieldId;
	uiType = AR_DATA_TYPE_CURRENCY;

	// Save the value of currency field
	Value = currencyStruct->value;
	Value += ARX_VALUE_SEPERATOR; // insert seperator char before code
	Value += (char*)currencyStruct->currencyCode; // save the currency code of the value

	// Convert and save the Timestamp as an integer value
	Value += ARX_VALUE_SEPERATOR; // insert seperator before Timestamp
	strWork.Format("%d", currencyStruct->conversionDate); 
	Value += strWork;

	// Convert and save the number of currency items contained in this structure
	Value += ARX_VALUE_SEPERATOR; // insert seperator before number of items
	strWork.Format("%d", currencyStruct->funcList.numItems); Value += strWork;

	// Save each value and currency code stored in the list structure
	// insert seperator in between and before each pair
	for(unsigned int i=0; 
		i<currencyStruct->funcList.numItems; 
		i++, pFuncCurrency++)
	{
		Value += ARX_VALUE_SEPERATOR;
		Value += pFuncCurrency->value;
		Value += ARX_VALUE_SEPERATOR;
		Value += pFuncCurrency->currencyCode;
	}

	return Value;
}

// This function will dump the text of the record to the File
CRecord::DumpARX(CARSConnection &arsConnect, CString Form, CStdioFile &File,
				 CString &strAttachDir, unsigned int *p_uiAttachNum, CString &strBuffer)
{
	POSITION posValue = GetHeadPosition(); // working value position
	CFieldValuePair FieldValue; // working reference

	// Insert the data row header
	strBuffer = ARX_HEADER_DATA;

	// Dump each field to the strBuffer
	for(int i=0; i<GetCount(); i++)
	{
		// dump the current field to the text file
		FieldValue = GetNext(posValue);
		FieldValue.DumpARX(arsConnect, Form, File, strAttachDir, 
						   p_uiAttachNum, EntryId, strBuffer);
	}	
	strBuffer += "\n"; // append end of line to record
}

CFieldValuePair::DumpARX(CARSConnection &arsConnect, CString Form, CStdioFile &File,
						CString &strAttachDir, unsigned int *p_uiAttachNum,
						CEntryId &EntryId, CString &strBuffer)
{
	// Insert a space before the value
//	File.Write(" ", 1);
	strBuffer += " ";

	CString cleanValue(Value);
	
	// dump the value to the file
	switch(uiType)
	{
	case AR_DATA_TYPE_NULL:
		//File.Write("\"\"", 2);
		strBuffer += "\"\"";
		break;
	// basically an value exported via this case will not have any double
		// quotes exported
	case AR_DATA_TYPE_INTEGER:
	case AR_DATA_TYPE_REAL:
	case AR_DATA_TYPE_ENUM:
	case AR_DATA_TYPE_TIME:
	case AR_DATA_TYPE_DECIMAL:
		//File.Write(LPCSTR(Value), Value.GetLength()); // write the value	
		strBuffer += Value;
		break;
		// any value exported via this case will have double quotes
		// inserted before and after the Value.
	case AR_DATA_TYPE_CHAR:
	case AR_DATA_TYPE_DIARY:
	case AR_DATA_TYPE_CURRENCY: // treat currency like a char because we need to double quotes
		cleanValue.Replace("\"", "\\\""); // insert a \ before each double quote
		cleanValue.Replace("\n", "\\r\\n"); // replace carriage returns with \r\n text
		//File.Write("\"", 1); // write double quote before value
		strBuffer += "\"";
		//File.Write(LPCSTR(cleanValue), cleanValue.GetLength()); // write the text value
		strBuffer += cleanValue;
		//File.Write("\"", 1); // write double quote after value
		strBuffer += "\"";
		break;
	case AR_DATA_TYPE_ATTACH:
		CString strExt; // working extenstion of the attachment
		CString strFile; // working file name with no extension
		CString strAttachNum; // working string for attachment number
		CString strFinishedAttach; // working completed attachment name to use when calling GetEntryBLOB()
		CString strFilePath; // the full path the the filename for saving the BLOB file

		// Save the file extension
		strExt = Value;
		strExt.MakeReverse();
		strExt = strExt.Left(4);
		strExt.MakeReverse();

		// Save the file name with out extenstion
		strFile = Value.Left(Value.GetLength() - 4);

		// convert the Attachment number to a string and increment it
		strAttachNum.Format("%d", *p_uiAttachNum);
		*p_uiAttachNum = *p_uiAttachNum + 1;

		// build the completed attachment name to save
//		strBuffer += "\"" + strAttachDir + "\\" + strFile + "_" + \
//			strAttachNum + strExt + "\"";
		strFinishedAttach = strAttachDir + "\\" + strFile + "_" + strAttachNum + strExt;
		strBuffer += "\"" + strFinishedAttach + "\"";

		// Lastly, save the attachment file BLOB
		// This function call needs the full path\filename to save the BLOB
		// this should be specified in strAttachFileName
		strFilePath = File.GetFilePath();
		strFilePath.Replace(File.GetFileName(), LPCSTR(strFinishedAttach)); // replace the filename with the relative attachment name
		DumpAttachment(arsConnect, Form, EntryId, uiFieldId, strFilePath);
		break;
	} // end switch
}

//DEL CString CFieldValuePair::CleanupQuotes(CString aString)
//DEL {
//DEL 	CString cleanString = aString;
//DEL 	int position = 0;
//DEL 	
//DEL 	while( (position = cleanString.Find("\"", position)) != -1)
//DEL 	{
//DEL 		cleanString.Insert(position, "\\"); // insert a backslash before the double quote
//DEL 	}
//DEL 
//DEL 	return cleanString;
//DEL }

////////////////////////////////////////////////////////////////////////////////
// Saves the file strFilePath after getting it from the ars server
CFieldValuePair::DumpAttachment(CARSConnection &arsConnect, CString Form, 
								CEntryId EntryId, ARInternalId arsFieldId,
								CString strFilePath)
{
	ARNameType arsFormName; // working variable
	AREntryIdList arsEntryId; // working variable
	ARLocStruct arsLocStruct; // working variable
	ARStatusList arsStatusList; // working variable

	// Build ARNameType formName
	strcpy( (char*)arsFormName, LPCSTR(Form.Left(sizeof(ARNameType))) ); // copy form name, truncating if necessary
	if(Form.GetLength() > sizeof(ARNameType))
	{
		// Output truncation to log file
		CString strLog("CFieldValuePair::DumpAttachment()");
		strLog += "\tForm name truncated due because size exceeded limits.  Original name: ";
		strLog += Form;
		strLog += INDENT;
		strLog += "Truncated name: ";
		strLog += (char*)arsFormName;
//		Log(strLog);
	}
	// Build AREntryIdList entryId. Will be entry id to get blob for
	if(!FillEntryId(EntryId, &arsEntryId))
	{
		// output to error log here
		CString strLog;
		strLog = "\tAttachment \"";
		strLog += strFilePath;
		strLog += "\" not saved because FillEntryId() failed.\t";
		strLog += "CFieldValuePair::DumpAttachment()";
		Log(strLog);
		ThrowARSException(LPCSTR(strLog), "CFieldValuePair::DumpAttachment()");
	}

	// Build ARLocStruct. Will hold the filename to save the attachment to.
	if(!FillARLocStruct(strFilePath, arsLocStruct))
	{
		// output to error log here
		CString strLog;
		strLog = "\tAttachment \"";
		strLog += strFilePath;
		strLog += "\" not saved because FillARLocStruct() failed.\t";
		strLog += "CFieldValuePair::DumpAttachment()";
		Log(strLog);
		ThrowARSException(LPCSTR(strLog), "CFieldValuePair::DumpAttachment()");
	}

	// Call ARGetEntryBLOB to dump the attachment
	if(ARGetEntryBLOB(&arsConnect.LoginInfo, arsFormName, &arsEntryId,
					  arsFieldId, &arsLocStruct, &arsStatusList) > AR_RETURN_OK)
	{
		ThrowARSException(arsStatusList, "CFieldValuePair::DumpAttachment");
	}

	// Free up heap memory
	FreeAREntryIdList(&arsEntryId, FALSE);
	FreeARLocStruct(&arsLocStruct, FALSE);
}

//////////////////////////////////////////////////////////////////////////////
// Fills up p_arsEntryId with the contents of EntryId
// IN: EntryId
// OUT: p_arsEntryId
// Returns: 0 - on error, 1 - on success
int CFieldValuePair::FillEntryId(CEntryId &EntryId, AREntryIdList *p_arsEntryId)
{
	// if EntryId is empty, throw an error
	if(EntryId.GetCount() == 0)
	{
		ThrowARSException(CREC_ENTRY_ID_EMPTY, "CFieldValuePair::FillEntryId()");
		return 0;
	}

	p_arsEntryId->numItems = EntryId.GetCount(); // store how many id's will make up this entry id list
	p_arsEntryId->entryIdList = (AREntryIdType*)malloc(sizeof(AREntryIdType) * EntryId.GetCount()); // allocated memory for entryId

	if(p_arsEntryId->entryIdList == NULL)	// ensure memory was allocated
	{
		p_arsEntryId->numItems = 0;
		AfxThrowMemoryException();
		return 0; // return empty entry id because of error
	}

	// now fill p_arsEntryId with the entry id's in the current CEntryId object
	AREntryIdType *pEntryId = p_arsEntryId->entryIdList;
	POSITION pos = EntryId.GetHeadPosition();
	for(unsigned int i=0;
		i<p_arsEntryId->numItems; 
		i++, pEntryId++)
	{
		// should make sure EntryId doesn't have any strings longer than pEntryId
		// can handle
		strcpy((char*)pEntryId, LPCSTR(EntryId.GetNext(pos)) );
	}
	
	// Don't call FreeAREntryIdList.  This should be called by the calling function
	return 1;
}

/////////////////////////////////////////////////////////////////////////////////////
// Fills arsLocStruct with contents of strFilePath.  Get's arsLocStruct ready
// to be used in call to ARGetEntryBLOB to dump attachment file to disk
// Returns: 0 - on error, 1 - on success
int CFieldValuePair::FillARLocStruct(CString &strFilePath, ARLocStruct &arsLocStruct)
{
	// allocate heap memory and error check
	char *pText = (char*)malloc(   sizeof(char) * (strFilePath.GetLength() + 1)  );
	if(!pText)
	{
		AfxThrowMemoryException();
		return 0;
	}

	// copy strFilePath into arsLocStruct
	strcpy(pText, LPCSTR(strFilePath));
	arsLocStruct.u.filename = pText;
	arsLocStruct.locType = AR_LOC_FILENAME;  // save type of ARLocStruct

	// Don't free heap memory for pText. The calling function needs to do 
	// that using FreeARLocStruct()
	return 1;
}


