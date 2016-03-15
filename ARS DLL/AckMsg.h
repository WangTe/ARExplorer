// AckMsg.h: interface for the CAckMsg class.
//
//////////////////////////////////////////////////////////////////////

#if !defined(AFX_ACKMSG_H__83AA38A7_1DBC_4E47_9CCB_E3B55D263FF2__INCLUDED_)
#define AFX_ACKMSG_H__83AA38A7_1DBC_4E47_9CCB_E3B55D263FF2__INCLUDED_

#if _MSC_VER > 1000
#pragma once
#endif // _MSC_VER > 1000

#define ACK_YES "YES"
#define ACK_NO "NO"
#define ACK_OK "OPERATION_OK"

class AFX_EXT_CLASS CAckMsg  : public CObject
{
	DECLARE_SERIAL(CAckMsg);
public:
	CAckMsg(const CAckMsg &Copy);
	unsigned int uiMsgCode;
	CAckMsg(LPCTSTR pStr, UINT uiType, UINT uiCode);
	void Serialize(CArchive &ar);
	enum msg_types { status = 0, // used for status' from function calls, not meant to be displayed
					 question,  // questions needed from the end user of the client
					 warning,  // warning, not critical, but should display warning dialog
					 error }; // something went wrong, display error dialog to user
	UINT uiMsgType;
	CString strMessage;
	CAckMsg();
	virtual ~CAckMsg();
	CAckMsg & operator =(const CAckMsg &newMsg);

};

#endif // !defined(AFX_ACKMSG_H__83AA38A7_1DBC_4E47_9CCB_E3B55D263FF2__INCLUDED_)
