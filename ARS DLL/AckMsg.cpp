// AckMsg.cpp: implementation of the CAckMsg class.
//
//////////////////////////////////////////////////////////////////////

#include "stdafx.h"
#include "AckMsg.h"

IMPLEMENT_SERIAL(CAckMsg, CObject, 1);

//////////////////////////////////////////////////////////////////////
// Construction/Destruction
//////////////////////////////////////////////////////////////////////

CAckMsg::CAckMsg()
{
	uiMsgType = status;
	strMessage = "none";
	uiMsgCode = 0;
}

CAckMsg::~CAckMsg()
{

}

void CAckMsg::Serialize(CArchive &ar)
{
	CObject::Serialize(ar);

	if(ar.IsStoring()) {
		ar << strMessage;
		ar << uiMsgType;
		ar << uiMsgCode;
	}else {
		ar >> strMessage;
		ar >> uiMsgType;
		ar >> uiMsgCode;
	}
}

CAckMsg & CAckMsg::operator =(const CAckMsg &newMsg)
{
	strMessage = newMsg.strMessage;
	uiMsgType = newMsg.uiMsgType;
	uiMsgCode = newMsg.uiMsgCode;

	return *this;
}

CAckMsg::CAckMsg(LPCTSTR pStr, UINT uiType, UINT uiCode)
{
	strMessage = pStr;
	uiMsgType = uiType;
	uiMsgCode = uiCode;
}

CAckMsg::CAckMsg(const CAckMsg &Copy)
{
	this->operator=(Copy);
}
