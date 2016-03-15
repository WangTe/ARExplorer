// CommPacket.h: interface for the CCommPack class.
//
//////////////////////////////////////////////////////////////////////

#if !defined(AFX_COMMPACK_H__10CCA9E4_F4A1_472B_87DA_EE78753DBC70__INCLUDED_)
#define AFX_COMMPACK_H__10CCA9E4_F4A1_472B_87DA_EE78753DBC70__INCLUDED_

#if _MSC_VER > 1000
#pragma once
#endif // _MSC_VER > 1000

#define MY_AR_MAX_NAME_SIZE 254


#define SIZE_DWORD sizeof(DWORD)
#define TIME_OUT		15 // default tcp/ip timeout

// Server port of backup server
#define SERVER_PORT 45090

// Empty control code
#define SERVICE_CONTROL_NONE	0

// Communication Packet Types
#define TCP_COMMAND_STOP			1
#define TCP_COMMAND_JOB				2
#define TCP_COMMAND_LIST_JOBS		3
#define TCP_COMMAND_LIST_ACTIVITY	4
#define TCP_COMMAND_DISABLE_JOB		5
#define TCP_COMMAND_ENABLE_JOB		6
#define TCP_COMMAND_GET_JOB			7
#define TCP_COMMAND_CHECK_DIR		8

// Communication Packet Class - to tell Service what type of packet was received
//struct CommPacket
//{
//	DWORD l_iCommType;
//	DWORD l_iControlCode;
//};

// Misc. sizes and max values
//#define SIZE_JOB_NAME	128 // max length of job name
//#define SIZE_JOB_TIME	16 // HHMMSSAMDDMMYYYY
//#define SIZE_JOB_NOTES	1024 // max length of job notes

// Communication Packet Class - Contains the client communication details
// (i.e. backup job name, backup job details, change/delete/create backup job, etc)

typedef BYTE type_FormName[MY_AR_MAX_NAME_SIZE + 1];

#endif // !defined(AFX_COMMPACK_H__10CCA9E4_F4A1_472B_87DA_EE78753DBC70__INCLUDED_)
