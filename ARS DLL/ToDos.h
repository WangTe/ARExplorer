///////////////////////////////////////////////////////////////////
// List all open to do's here
//
//     Done:	Modify CFieldValuePair::StoreCurrency() so it inserts 
//				seperator char into Value.  You need to find out what 
//				the correct seperator character is ASCII CODE: 4.
//
// Not Done:	Create object to hold attachment data (BLOB).  Call 
//				it CAttachmentField
//
//     Done:	Make sure you insert a '\' before any double quote
//				in any Char or Diary field before you export the text
//				to the ARExport file.
//
//     Done:	Make sure the embeded \n in char and diary fields
//				will be converted properly to a \r\n in the text file
//				using File.Write() fuction.
//
//     Done:	Modify CFieldValuePair::DumpARX in the Attachment section
//				so Unix files attached that have no file extension will be saved properly.
//				Right now, i'm blindly inserting an index before the 4th to last char
//				in the filename.  I.E. if the filename is "Some Doc.doc" it will be saved as
//				"Some Doc_#.doc"
//
//	    Done:	Finish CAttachment::DumpARX() to call ARGetEntryBLOB().  Needs
//				to populate AREntryIdList and ARInternalId first.
//
//  Not Done:	Create a macro function that takes a function name, text, etc
//				and concats it into one string for passing the Log()
//
//      Done:	Add error logging to CFieldValuePair::DumpAttachment()
//
//  Not Done:	Add code globally which allows the Backup directory to have a the last
//				char be a \
//
//  Not Done:	Add the rest of code to CDumpARS class which will export the 
//				form name, data types, field ids and DATA record headers 
//				to the .arx file. (This is in progress)
//
//     Done:	Change CFieldValuePair::DumpARX so it populates a string instead
//				of writing to the file immediately.  Then in CRecordList::DumpARX
//				write the record after the entire string is populated.  This 
//				way you can dump bad records (i.e. unkown data types, etc) to a 
//				log file and continue the backup process.
//
int ToDo;

//////////////////////////////////////////////////////////////////////////
// List of open bugs
//------------------------------------------------------------------------
//     Fixed:	No data is being exported, only double quotes are being
//				written to the file.
//
//     Fixed:	There is an exponential loop.  I.E. - 28 records in the 
//				Group form will dump 28 lines, each line having 28 more 
//				fields exported.  Also it will one set of 28 lines for
//				each record found.
//
//     Fixed:	Find out why my DumpARX is exporting a 32 char right after the DATA header
//	Soltuion:	When I replace File.Write in all of the DumpARX() functions with a
//				strBuffer (string buffer) and then used File.Write() once
//				to write the entire record, it fixed it.
//
// Not Fixed:	CRecordList::GetRecord isn't populating the EntryId when 
//				the record is returned.


//////////////////////////////////////////////////////////////////////////
// List of things that still need to be done
// Not Done:	Test DumpARX() to see if it dumps attachment names into
//				the ARX file correctly.
// Not Done:	Test DumpARX() to see if it save the attachements into
//				the correct sub-directory.
//
// Not Done:	Finish the CARNameType class to return an ARNameType*
//				from operator ARNTP()
//
// Not Done:	Finish CFieldList class,