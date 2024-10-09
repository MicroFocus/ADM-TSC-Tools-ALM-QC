Deactivate Projects using SA API

1.  Install SA client from "Register HP AlM Site Administration" site of
    ALM instance.

2.  Use command line to execute.

Usage:

c:\\windows\\syswow64\\cscript DeactivateProject.vbs \<serverurl\>
\<username\> \<userpassword\> \<workingdirectory\> \<projectlistfile\>

Example:

c:\\windows\\syswow64\\cscript DeactivateProject.vbs
https://\<server\>/qcbin myname \##### c:\\temp ProjectList.txt

Notes: ProjectList.txt should exist under the working directory (Ex.
C:\\temp). Each project name should have the format of
domain_name.project_name (Ex. Default.TestDirector_Demo). Each project
should be placed under an individual line.
