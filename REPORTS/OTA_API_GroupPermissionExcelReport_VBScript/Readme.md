Deactivate Projects using SA API

1.  Install OTA client from "Register ALM" site of ALM instance.
	Example:
	
	32-bit
	https://<server>/qcbin/start_a.jsp?common=true
	
	64-bit
	https://<server>/qcbin/start_a.jsp?common=true&comsurrogate=true

2.  Use command line to execute.

Usage:

32-bit
c:\\windows\\syswow64\\cscript GroupPermissionExcel.vbs https://<almhost>/qcbin/ <user> <password> <path>

64-bit
cscript GroupPermissionExcel.vbs https://<almhost>/qcbin/ <user> <password> <path>

Example:

c:\windows\syswow64\cscript GroupPermissionExcel.vbs https://almserver.saas.microfocus.com/qcbin user password c:\temp


