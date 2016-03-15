
#include "StdAfx.h"

#include "Schedule.h"

#include "BackupJob.h"
#include "Globals.h"

//////////////////////////////////////////////////////////////////////
// Construction/Destruction
//////////////////////////////////////////////////////////////////////
// CBackupJob impelementation
IMPLEMENT_SERIAL(CBackupJob, CObject, 1);

void CBackupJob::Serialize(CArchive &archive)
{
	CObject::Serialize( archive );

	if(archive.IsStoring()) { // Serialize
		archive << strBackupDir;
		archive << strJobName;
		archive << strNotes;
		archive << strQualifier;
		archive << uiMaxRecords;
		archive << uiClientAction;
		archive << bEnabled;
		archive << m_tLastRunTime;
		archive << bIncremental;
	}
	else { // De-serialize
		archive >> strBackupDir;
		archive >> strJobName;
		archive >> strNotes;
		archive >> strQualifier;
		archive >> uiMaxRecords;
		archive >> uiClientAction;
		archive >> bEnabled;
		archive >> m_tLastRunTime;
		archive >> bIncremental;
	}
	ars_LoginInfo.Serialize( archive );
	Schedule.Serialize(archive);

	if(ars_Forms.IsSerializable()) {
		ars_Forms.Serialize( archive );
	log.Log(CFileLogging::level_lower, "CBackupJob:Serialize(): Done Serializing list of forms.", NULL);
	}
	log.Log(CFileLogging::level_lower, "CBackupJob:Serialize(): Done serializing job: ", LPCTSTR(strJobName), NULL);
}

CBackupJob & CBackupJob::operator =(CBackupJob &newJob)
{
	ars_Forms = newJob.ars_Forms;
	ars_LoginInfo = newJob.ars_LoginInfo;
	Schedule = newJob.Schedule;
//	IncSchedule = newJob.IncSchedule;
	strBackupDir = newJob.strBackupDir;
	strJobName = newJob.strJobName;
	strNotes = newJob.strNotes;
	strQualifier = newJob.strQualifier;
	uiMaxRecords = newJob.uiMaxRecords;
	uiClientAction = newJob.uiClientAction;
	bEnabled = newJob.bEnabled;
	m_tLastRunTime = newJob.m_tLastRunTime;
	bIncremental = newJob.bIncremental;

	return *this;
}

CBackupJob::CBackupJob()
{
	m_tLastRunTime = 0;
	bIncremental = false;
	uiMaxRecords = 0;
}

CBackupJob::~CBackupJob()
{

}

CBackupJob::CBackupJob(CBackupJob &Copy)
{
	operator=(Copy);
}

////////////////////////////////////////////////////////////////////
// Description: Gets the job type from the attached schedule
//				and returns it
//		once,
//		daily,
//		weekly,
//		monthly,
//		now
UINT CBackupJob::GetType()
{
	return Schedule.GetType();
}

CTime CBackupJob::GetStartTime()
{
	return Schedule.GetStartTime();
}

BOOL CBackupJob::GetEnabledFlag()
{
	return bEnabled;
}

CTime CBackupJob::GetLastRunTime()
{
	return m_tLastRunTime;
}

CTime & CBackupJob::SetStartTime(CTime &m_newTime)
{
	return Schedule.SetStartTime(m_newTime);
}

void CBackupJob::SetLastRunTime(CTime &newTime)
{
	m_tLastRunTime = newTime;
}

BOOL CBackupJob::GetIncFlag()
{
	return bIncremental;
}
