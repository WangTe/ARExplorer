#ifndef __BACKUP_JOB_H
#define __BACKUP_JOB_H


#if _MSC_VER > 1000
#pragma once
#endif // _MSC_VER > 1000

#include "ARSConnection.h"
#include "Schedule.h"
#include "FormList.h"


class AFX_EXT_CLASS CBackupJob : public CWinThread
{
	DECLARE_SERIAL(CBackupJob);
public:
	CBackupJob();
	virtual ~CBackupJob();
	enum job_operations { none = 0, create_job, modify_job, delete_job, run_job, status_job}; 
	void Serialize( CArchive &archive );
	enum status { stopped, suspended, scheduled, running  }; // job status'
	CBackupJob & operator =(CBackupJob &newJob);

// public data members
public: 
	BOOL GetIncFlag();
	void SetLastRunTime(CTime &newTime);
	CTime & SetStartTime(CTime &m_newTime);
	CTime GetLastRunTime();
	BOOL GetEnabledFlag();
	CTime GetStartTime();
	UINT GetType();
	BOOL bEnabled;
	CBackupJob(CBackupJob &Copy);
	UINT uiClientAction;
	UINT uiMaxRecords;
	CString strQualifier;
//	unsigned int uiOverwriteFlag; // not really sure what this is for, maybe delete it
	CString strBackupDir;
	CFormList ars_Forms; // list of forms for backup job
	CString strNotes;
	CARSConnection ars_LoginInfo; // login, password, & server name
	CSchedule Schedule; // Complete Backup schedule
//	CSchedule IncSchedule; // Incremental Schedule (for appending dumps)
	CString strJobName;
protected:
	BOOL bIncremental; // true if it's an incremental backup job, false if complete backup job
	CTime m_tLastRunTime;
};


#endif