1.  Register the OTA Client via installing the Client Registration

Example:

32-bit

![A screenshot of a computer AI-generated content may be
incorrect.](media/image1.png){width="4.013438320209974in"
height="2.3934853455818024in"}

64-bit

![A screenshot of a computer AI-generated content may be
incorrect.](media/image2.png){width="4.004764873140857in"
height="2.3917344706911634in"}

Notes:

-   Admin permission is needed to register the ALM Client.

-   Bit version should match the bit version of your Excel

2.  Launch the RunReport.xlsm file.

3.  Navigate to the Connection tab and click on Execute button. The SaaS
    Excel Reporting Dialog window appears.

4.  Enable macro if prompted.

Note:

-   Authentication using either username/password or apikey/apisecret.

-   Enable macro if prompted

5.  Enter the ALM URL and click on "Initialize Server".

6.  Enter your username/password or apikey/secret.

7.  Select your domain and project.

8.  Under "Select Task", select "Run Report".

9.  Specified the location of your log file. Attachments and log files
    will be generated under this directory.

![A screenshot of a computer AI-generated content may be
incorrect.](media/image3.png){width="2.9856233595800523in"
height="2.499115266841645in"}

10. Click Run. Run Report dialog window appears.

11. Enter your Execution Start and End Date you want to filter by. The
    date here is the RN_EXECUTION_DATE.

![A screenshot of a computer AI-generated content may be
incorrect.](media/image4.png){width="6.5in"
height="2.209722222222222in"}

12. Click on Run button to execute the report.

Note: If you have a large amount of data, then please make use of the
Execution Date filter to limit the output.
