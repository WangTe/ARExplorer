// Form.h: interface for the CForm class.
//
//////////////////////////////////////////////////////////////////////

#if !defined(AFX_FORM_H__EBCFA63A_2DBC_475C_92E3_C0AEBFD75922__INCLUDED_)
#define AFX_FORM_H__EBCFA63A_2DBC_475C_92E3_C0AEBFD75922__INCLUDED_

#if _MSC_VER > 1000
#pragma once
#endif // _MSC_VER > 1000

#include <afxtempl.h>
#include "ARSConnection.h"

////////////////////////////////////////////////////////////////////////
// Error messages for this class
const char FORM_NAME_EMPTY[] = "The form name was empty.";

////////////////////////////////////////////////////////////////////////
// Type Defs for returning ARS types from classes
typedef ARInternalId* ARIIP;
typedef ARInternalIdList* ARIILP;

////////////////////////////////////////////////////////////////////////
// Object to hold single field details (name, id, type)
class AFX_EXT_CLASS CField  
{
public:
	CString TypeText;
	CField(const CField &Copy);
	unsigned int Type;
	ARInternalId Id;
	CString Name;
	CField();
	virtual ~CField();
	CField operator =(const CField &newField);
	operator ARIIP();
};

////////////////////////////////////////////////////////////////////////
// object to hold list of Remedy fields.  Also contains
// easy conversion function to ARInternalIdList for ARS function calls
//
// This class is pretty much a redone copy of CFieldList in the 
// AR_Explorer_API project.  This class will be reworked into a true
// object oriented class.
class AFX_EXT_CLASS CFieldList : public CList <ARInternalId, ARInternalId &>
{
public:
	ARInternalIdList arsIdList;
	CField GetNextField(CARSConnection &arsConnect, POSITION &pos);
	CFieldList(CString newForm);
	CString Form;
	FillFieldList(CARSConnection &arsConnect, unsigned long ulFieldType);
	FillFieldList(CARSConnection &arsConnect, CString newForm, 
				  unsigned long ulFieldType);
	CFieldList(LPCTSTR pForm);
	virtual ~CFieldList();
	operator ARIILP();

};

///////////////////////////////////////////////////////////////////////
// Object to hold details about a single form.
class AFX_EXT_CLASS CForm  
{
public:
	CString LastModifiedBy;
	CTime Timestamp;
	CString HelpText;
	CString Type;
	CString Name;
	CForm();
	virtual ~CForm();

protected:
};



#endif // !defined(AFX_FORM_H__EBCFA63A_2DBC_475C_92E3_C0AEBFD75922__INCLUDED_)
