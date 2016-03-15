// Form.cpp: implementation of the CForm class.
//
//////////////////////////////////////////////////////////////////////

#include "stdafx.h"
#include "Form.h"

#ifdef _DEBUG
#undef THIS_FILE
static char THIS_FILE[]=__FILE__;
#define new DEBUG_NEW
#endif

#include "ARSException.h"

//////////////////////////////////////////////////////////////////////
// Construction/Destruction
//////////////////////////////////////////////////////////////////////

CForm::CForm()
{
	Timestamp = 0;

}

CForm::~CForm()
{

}


//DEL CForm::CForm(CString &newName, int newType, CString &newHelpText, ARTimestamp newTimestamp)
//DEL {
//DEL 	Name = newName;
//DEL 	Type = newType;
//DEL 	HelpText = newHelpText;
//DEL 	Timestamp = newTimestamp;
//DEL }

CFieldList::FillFieldList(CARSConnection &arsConnect, CString newForm, 
						  unsigned long ulFieldType = AR_FIELD_TYPE_DATA)
{
	// save the new form name and ensure proper size of Form Name
	Form = newForm;
	// call FillFieldList
	FillFieldList(arsConnect, ulFieldType);
}

CFieldList::FillFieldList(CARSConnection &arsConnect, unsigned long ulFieldType = AR_FIELD_TYPE_DATA)
{
	ARStatusList arsStatusList; // working status list
	CField Field; // working CField object

	// Reset the field data.
	RemoveAll();

	// If Form is empty, throw error
	if(Form.IsEmpty())
		ThrowARSException(FORM_NAME_EMPTY, "CFieldList::FillFieldList()");

	// Make sure Form name isn't too long
	if(Form.GetLength() > sizeof(ARNameType))
		Form = (Form.Left(sizeof(ARNameType) - 1));

	// Get the list of fields and store them in arsIdList.
	if( ARGetListField((ARCSP)arsConnect,
								  (ARNT)(LPCTSTR)Form,
								  ulFieldType,
								  0,
								  &arsIdList,
								  &arsStatusList ) > AR_RETURN_OK)
	{
		FreeARInternalIdList(&arsIdList, FALSE); // free the working id list
		ThrowARSException(arsStatusList, "CFieldList::FillFieldList()");
	}

	// Add each field id to CFieldList
	for(unsigned int i=0; i<arsIdList.numItems; i++)
		AddTail(arsIdList.internalIdList[i]);

	// Free heap memory
	FreeARStatusList(&arsStatusList, FALSE);
}

CFieldList::CFieldList(CString newForm)
{
	Form = newForm;
}

CField CFieldList::GetNextField(CARSConnection &arsConnect, POSITION &pos)
{
	CField workingField; // working field object
	ARInternalId arsId; // working Field id
	ARNameType arsFieldName; // working field name
	unsigned int arsDataType; // working field data type
	ARStatusList arsStatusList; // working status list

	arsId = GetNext(pos); // get's next id in the list and save in working variable

	// Retrieve field name for current field Id
	if(ARGetField((ARCSP)arsConnect, // IN: login info
							(ARNT)(LPCTSTR)Form,	// IN: Form name to get field from
							arsId,	// IN: Field Id to get
							arsFieldName,	// OUT: Field Name
							NULL,							
							&arsDataType, // OUT: Field Data Type
							NULL,
							NULL,
							NULL,
							NULL,
							NULL,
							NULL,
							NULL,
							NULL,
							NULL,
							NULL,
							NULL,
							&arsStatusList ) > AR_RETURN_OK)
	{
		ThrowARSException(arsStatusList, "CFieldList::GetNext()");
	}

	// save the gotten values into Field
	workingField.Id = arsId;
	workingField.Name = arsFieldName;
	workingField.Type = arsDataType;

	// Save Data Type as Text Values
	switch(arsDataType)
	{
	case AR_DATA_TYPE_CHAR:
		workingField.TypeText = "CHAR";
		break;
	case AR_DATA_TYPE_ENUM:
		workingField.TypeText = "ENUM";
		break;
	case AR_DATA_TYPE_INTEGER:
		workingField.TypeText = "INTEGER";
		break;
	case AR_DATA_TYPE_REAL:
		workingField.TypeText = "REAL";
		break;
	case AR_DATA_TYPE_DIARY:
		workingField.TypeText = "DIARY";
		break;
	case AR_DATA_TYPE_TIME:
		workingField.TypeText = "TIME";
		break;
	case AR_DATA_TYPE_BITMASK:
		workingField.TypeText = "BITMASK";
		break;
	case AR_DATA_TYPE_BYTES:
		workingField.TypeText = "BYTES";
		break;
	case AR_DATA_TYPE_DECIMAL:
		workingField.TypeText = "DECIMAL";
		break;
	case AR_DATA_TYPE_ATTACH:
		workingField.TypeText = "ATTACH";
		break;
	case AR_DATA_TYPE_CURRENCY:
		workingField.TypeText = "CURRENCY";
		break;
	case AR_DATA_TYPE_DATE:
		workingField.TypeText = "DATE";
		break;
	case AR_DATA_TYPE_TIME_OF_DAY:
		workingField.TypeText = "TIMEOFDAY";
		break;
	default:
		workingField.TypeText = "UNKNOWN";
		break;
	} // end switch

	// return the field
	return workingField;
}

CFieldList::operator ARIILP()
{
	return &arsIdList;
}
