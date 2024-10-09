## Test Instance purge tool

Example tool for purging runs from the Test Instance level. This is useful when purging larger Test Sets.
Once the script completes, it will notify the ALM user via email.

This tool is meant to run in the command line using the ALM QC OTA library.

1. Ensure the OTA Library is registered on your machine
2. Copy the PURGE.VBS file to a local folder
3. Open the VBS file in notepad or your preferred editing tool and update the following parameters:
```
purgeName = "Email Subject"
purgeUser = "QC User for email notification"
purgeKey = "<QC API Key>"
purgeSecret = "<QC API Secret>"
serverURL = "<QC Server URL>"
domainName = "<QC Domain name>"
project Name = "<QC Project name>"
```

**NOTE:** The tool is set to keep the last 5 runs for each test instance. You can change this within the PurgeRun2 method.
More information on the PurgeRun2 method: https://admhelp.microfocus.com/alm/api_refs/ota/Content/ota/topic9008.html

4. Prepare your SQL query, the tool expects the Test Set ID in the first column and the Test Instance ID in the second column.
Example query to get Test Instances for Test Set ID 123:
```
SELECT RN_CYCLE_ID, RN_TESTCYCL_ID FROM RUN WHERE (RN_CYCLE_ID IN(123)) GROUP BY RN_CYCLE_ID, RN_TESTCYCL_ID
```
5. Start the Command Prompt window in 32 bit mode.
If you are on a 64 bit machine use:
```
Start -> Run -> %windir%\SysWoW64\cmd.exe 
```
7. Change to the folder directory. For example:
```
CD C:\PurgeFolder
```
9. Run the script with the cscript command:
```
CScript PURGE.VBS "<SQL>"
```
For example, to run with the SQL example for Test Set ID 123 use:
```
Cscript PURGE.VBS "SELECT RN_CYCLE_ID, RN_TESTCYCL_ID FROM RUN WHERE (RN_CYCLE_ID IN(123)) GROUP BY RN_CYCLE_ID, RN_TESTCYCL_ID"
```
**HOT TIP:**
It's possible to purge multiple Test Sets by editing the SQL Query, for example:
```
SELECT RN_CYCLE_ID, RN_TESTCYCL_ID FROM RUN WHERE (RN_CYCLE_ID IN(123, 456)) GROUP BY RN_CYCLE_ID, RN_TESTCYCL_ID
```
