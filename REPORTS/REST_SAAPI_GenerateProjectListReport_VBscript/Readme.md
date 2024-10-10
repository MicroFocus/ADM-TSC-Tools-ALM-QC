Generate list of projects using SA API:

1.  Install SA client from "Register HP AlM Site Administration" site of
    ALM instance.

2.  Use command line to execute.

Usage:

32-bit

c:\\windows\\syswow64\\cscript GenerateProjectList.vbs \<serverurl\>
\<username\> \<userpassword\> \<workingdirectory\> \<projectlistfile\>

64-bit

cscript GenerateProjectList.vbs \<serverurl\> \<username\>
\<userpassword\> \<workingdirectory\> \<projectlistfile\>

Example:

c:\\windows\\syswow64\\cscript GenerateProjectList.vbs
https://\<server\>/qcbin myname \##### c:\\temp ProjectList.txt

Notes: The ProjectList.txt will be created under your working directory
(Example: c:\\temp).
