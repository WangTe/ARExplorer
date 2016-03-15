// Field.cpp: implementation of the CField class.
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

CField::CField()
{
	Id = 0;
	Name.Empty();
	Type = 0;
}

CField::~CField()
{

}

// Assignment operator
CField CField::operator =(const CField &newField)
{
	Id = newField.Id;
	Name = newField.Name;
	Type = newField.Type;
	TypeText = newField.TypeText;

	return *this;
}

// Copy constructor
CField::CField(const CField &Copy)
{
	// call the assignment operator
	Id = Copy.Id;
	Name = Copy.Name;
	Type = Copy.Type;
	TypeText = Copy.TypeText;
}

// Conversion to an ARInternalId *
CField::operator ARIIP()
{
	return &Id;
}

