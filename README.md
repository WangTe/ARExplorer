# ARExplorer

This project used to be closed source, but I decided to post it recently hoping some Remedy developers might find this code useful.

AR Explorer is a tool, based on BMC Software's Remedy AR System (tm).  The AR Explorer API leverages MFC 4.2 and the AR System API version 4.x.

The AR Explorer API is implemented as a COM DLL.  And the GUI was built using Visual Basic.  

The tools purpose is to enable Remedy Developers to search the often complex and convoluded "workflows" (a.k.a. code) of the AR System, so they can find and understand Remedy code more quickly.  It also has capabilities to mass-update certain blocks of code.  For example, some fields are shared across hundreds of forms, like 'Submitted By'.  AR Explorer can search all fields in the entire system quickly, and then modify attributes of the 'Submitted By' fields so they're consistent across the entire system.  Doing so using BMC's tools could take hours or even days of effort, and bits of code would most likely be missed.  But, using AR Explorer, performing the same modification takes a few minutes and is automated.

