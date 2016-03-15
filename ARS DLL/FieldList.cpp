// FieldList.cpp: implementation of the CFieldList class.
//
//////////////////////////////////////////////////////////////////////

#include "stdafx.h"
#include "Form.h"

#ifdef _DEBUG
#undef THIS_FILE
static char THIS_FILE[]=__FILE__;
#define new DEBUG_NEW
#endif

//////////////////////////////////////////////////////////////////////
// Construction/Destruction
//////////////////////////////////////////////////////////////////////

CFieldList::CFieldList(LPCTSTR pForm = NULL)
{
	Form = pForm;
	arsIdList.numItems = 0;
	arsIdList.internalIdList = NULL;
}

CFieldList::~CFieldList()
{
	if(arsIdList.internalIdList != NULL)
		FreeARInternalIdList(&arsIdList, FALSE);
}
