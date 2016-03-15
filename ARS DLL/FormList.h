// FormList.h: interface for the CFormList class.
//
//////////////////////////////////////////////////////////////////////

#if !defined(AFX_FORMLIST_H__AA86E767_DD9E_4AF9_8F68_FDAA2FC72A4C__INCLUDED_)
#define AFX_FORMLIST_H__AA86E767_DD9E_4AF9_8F68_FDAA2FC72A4C__INCLUDED_

#if _MSC_VER > 1000
#pragma once
#endif // _MSC_VER > 1000

#include "Form.h"
#include "ARSConnection.h"
#include "StdAfx.h"

/////////////////////////////////////////////////////////////////////
// CFormList Custom Error Messages
#define FL_ERR_ITEM_NOT_FOUND "Item not found in list."
#define FL_ERR_LIST_EMPTY "Couldn't get an item, because form list was empty."

class AFX_EXT_CLASS CFormList : public CStringList 
{
	DECLARE_SERIAL(CFormList);
public:
	CForm Item(CARSConnection &arsConnection, int iIndex); // this array is zero indexed
	void GetForms(CARSConnection &arsConnection, unsigned int uiType);
	CFormList();
	virtual ~CFormList();
	CFormList & operator =(CFormList &newList);
};

#endif // !defined(AFX_FORMLIST_H__AA86E767_DD9E_4AF9_8F68_FDAA2FC72A4C__INCLUDED_)
