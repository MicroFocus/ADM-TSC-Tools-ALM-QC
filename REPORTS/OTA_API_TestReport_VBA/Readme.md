Export Tests Macro:

Requirements:

1.  This example uses OTA. OTA is a DCOM object that needs to be install
    on the machine. You can install OTA by either running the
    TDConnectivity or ALM Client installation.

TDConectivity:

ALM Client Registration (click on Register ALM link):

2.  Excel 32-bit only

3.  Excel 64-bit requires following the steps discussed in the link
    below:

<https://community.microfocus.com/t5/ALM-QC-User-Discussions/Workaround-OTA-amp-64-Bit-Applications-Excel-Word-etc/m-p/1639028/thread-id/97290/highlight/true>

4.  API Key for SSO authentication. If you don't have an API Key , then
    you can request your Administrator to create one.

![Graphical user interface, application Description automatically
generated](./media/media/image1.png){width="3.6881397637795277in"
height="3.1799048556430445in"}

Usage:

1.  Run the TestReport.xlsm. You should now see a Connection tab.

![Graphical user interface, application, table Description automatically
generated with medium
confidence](./media/media/image2.png){width="5.9375in"
height="1.78125in"}

Note: You will need to enable macro if prompted.

2.  Click on "Execute" button.

![Graphical user interface, application, table, Excel Description
automatically
generated](./media/media/image3.png){width="4.802083333333333in"
height="4.229166666666667in"}

3.  Enter the Server URL and click on "Initialize Server".

![Graphical user interface, application Description automatically
generated](./media/media/image4.png){width="6.020833333333333in"
height="5.020833333333333in"}

4.  Enter your user name and password. Click on "Authenticate".

![Graphical user interface, application Description automatically
generated](./media/media/image5.png){width="5.989583333333333in"
height="5.041666666666667in"}

Notes: For SSO, you need to enable the "APIKey" checkbox and enter your
API Key and Secret.

5.  Select your domain(s) and click on "Select Domain(s)" button.

6.  Select your project(s) and click on "Select Projects" button.

![Graphical user interface, application Description automatically
generated](./media/media/image6.png){width="6.03125in"
height="5.083333333333333in"}

Notes: Hold down shift key to multi-select items from list or click on
checkbox to select all.

7.  From the "Select Task" dropdown, select "Set Configuration". Click
    on "Run" button.

![Graphical user interface, application Description automatically
generated](./media/media/image7.png){width="5.989583333333333in"
height="5.052083333333333in"}

8.  Make your fields selection. Use the buttons to move the fields from
    "Available" list to "Selected" list. You can move a field up and
    down by clicking on corresponding buttons. Click on "Save".

![Graphical user interface, text, application Description automatically
generated](./media/media/image8.png){width="5.17422353455818in"
height="5.15321741032371in"}

9.  The fields you selected will be saved to the "Configuration" tab.
    You do not need to run this again for subsequent reports if no
    changes are needed. You can manually edit the Configuration tab to
    do field selection.

![Table Description automatically
generated](./media/media/image9.png){width="5.697916666666667in"
height="2.8125in"}

10. Select "Test Report" task. Click on "Run" button.

![Graphical user interface, application Description automatically
generated](./media/media/image10.png){width="6.020833333333333in"
height="5.083333333333333in"}

11. Report default to all folders under Subject. You can edit the
    Subject Path to filter.

![Graphical user interface, text, application Description automatically
generated](./media/media/image11.png){width="6.186811023622047in"
height="3.6979166666666665in"}

Note: Accept default Subject\\ to generate report for all folders below
the Subject node.

12. Click on Run button to generate the report.

13. Log entries will be generated and saved to c:\\temp.

![Graphical user interface, text, application Description automatically
generated](./media/media/image12.png){width="6.5in"
height="4.277777777777778in"}

14. Upon completion, a new tab with the project name is created. The
    report with the fields selected will be displayed in this tab.

![Graphical user interface, application, table, Excel Description
automatically generated](./media/media/image13.png){width="6.5in"
height="2.8569444444444443in"}
