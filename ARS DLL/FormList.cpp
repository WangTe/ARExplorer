// FormList.cpp: implementation of the CFormList class.
//
//////////////////////////////////////////////////////////////////////

#include "stdafx.h"
#include "FormList.h"
#include "ARSException.h"

#ifdef _DEBUG
#undef THIS_FILE
static char THIS_FILE[]=__FILE__;
#define new DEBUG_NEW
#endif

//////////////////////////////////////////////////////////////////////
// Construction/Destruction
//////////////////////////////////////////////////////////////////////

IMPLEMENT_SERIAL(CFormList, CStringList, 1);


CFormList::CFormList()
{
}

CFormList::~CFormList()
{

}

/////////////////////////////////////////////////////////////////
// Returns a CForm object of the form specified by iIndex
// iIndex is zero index for the list item to get
CForm CFormList::Item(CARSConnection &arsConnection, int iIndex)
{
	CForm Form;  // form to return

	// error if list is empty
	if(GetCount() == 0)
		ThrowARSException(FL_ERR_LIST_EMPTY, "CFormList::Item()");

	// error if iIndex is too big (i.e. past end of list)
	if(iIndex > (GetCount() - 1))
		ThrowARSException("iIndex is too large for list.", "CFormList::Item()");

	// Temp variables to hold return values from AR Server
	POSITION pos = FindIndex(iIndex); // working position for form name
	ARStatusList statusList;
	ARCompoundSchema compoundSchema;
	char *pHelpText; // temp ptr to hold help text
	ARAccessNameType lastModifiedBy; // temp to hold last modified by
	ARTimestamp modifiedTime; // temp to hold last mod time
	ARNameType arFormName; // temp to hold form name
//	CString strFormName; // temp to hold form name
//	strFormName = operator[](iIndex).Left(sizeof(ARNameType)); // shorten the form name if it's too long, i.e. if CFormList was populated directly instead of through GetForms()
	strcpy(arFormName, LPCSTR(GetAt(pos).Left(sizeof(ARNameType))));// convert to ARNameType format for function call
	
	// Get form props from server
	if(ARGetSchema(&arsConnection.LoginInfo, //INPUT: login info
					arFormName, //INPUT: name of form to get
					&compoundSchema, //OUT: will hold the type of form
					NULL, //OUT: group list
					NULL, //OUT: admin group list
					NULL, //OUT: query list fields
					NULL, //OUT: sort list
					NULL, //OUT: form index list
					NULL, //OUT: defautl view
					&pHelpText, //OUT: help text
					&modifiedTime, //OUT: last modified time
					NULL, //OUT: form owner
					lastModifiedBy, //OUT: who modified form last
					NULL, //OUT: change history
					NULL, //OUT: object properties for source code control
					&statusList) > AR_RETURN_OK)
	{
		ThrowARSException(statusList, "CFormList::Item()"); // error, throw exception
	}
	FreeARStatusList(&statusList, FALSE);

	try
	{
		// save the form type
		switch(compoundSchema.schemaType) 
		{
		case AR_SCHEMA_REGULAR:
			Form.Type = "Regular";
			break;
		case AR_SCHEMA_JOIN: 
			Form.Type = "Join";
			break;
		case AR_SCHEMA_VIEW:
			Form.Type = "View";
			break;
		case AR_SCHEMA_DIALOG:
			Form.Type = "Display";
			break;
		case AR_SCHEMA_VENDOR:
			Form.Type = "Vendor";
			break;
		default:
			Form.Type = "Unsupported";
			break;
		}// end switch()

		// Save the form mod time
		Form.Timestamp = modifiedTime;

		// Save the form mod by
		Form.LastModifiedBy = lastModifiedBy;

		// save the help text
		Form.HelpText = pHelpText;

		// save the form name
		Form.Name = arFormName;
	}
	catch(CMemoryException *)
	{
//		FreeARStatusList(&statusList, FALSE);
		FreeARCompoundSchema(&compoundSchema, FALSE);
		free ( pHelpText );
		throw;				
	}

	FreeARCompoundSchema(&compoundSchema, FALSE);
	free(pHelpText);



	return Form;
}

/////////////////////////////////////////////////////////////////////
// Fills the list with all the form names on the server
void CFormList::GetForms(CARSConnection &arsConnection, unsigned int uiType)
{
	ARStatusList statusList;
	ARNameList nameList;	// list of form names
	ARNameType *pNameType; // pointer to a single form name
	unsigned int i = 0;

	// get the list of form names
	if(ARGetListSchema(&arsConnection.LoginInfo,
					   0, // change since
					   uiType, // form type to get
					   NULL, // name of uplink/downlink form
					   NULL, // id list of fields to qualify forms retrieved
					   &nameList, // return value, the list of form names
					   &statusList) > AR_RETURN_OK)
	{
		FreeARNameList(&nameList, FALSE);
		ThrowARSException(statusList, "CFormList::GetForms()");
	}
	FreeARStatusList(&statusList, FALSE);

	CString strFormName;
	for(i=0, pNameType = nameList.nameList; 
		i < nameList.numItems; 
		i++, pNameType++)
	{
		strFormName = (char*)pNameType;
		AddTail(strFormName);
	}

	// Free heap memory
	FreeARNameList(&nameList, FALSE);
}


//DEL void CFormList::Serialize(CArchive &archive)
//DEL {
//DEL 
//DEL }

CFormList & CFormList::operator =(CFormList &newList)
{
	RemoveAll();
	AddTail(&newList);

	return *this;
}
