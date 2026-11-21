# PIM Group Manager
 
**PIM Group Manager** is a PowerShell-based graphical tool for creating Microsoft Entra groups.
 
Although the tool runs through PowerShell, most group creation and configuration tasks are completed through the graphical interface. You do not need to manually run Microsoft Graph commands to create groups.
 
The tool can create:
 
- PIM-enabled, role-assignable security groups
- Standard security groups
- Dynamic security groups
- Microsoft 365 groups
- Dynamic Microsoft 365 groups
- Multiple groups at once using a CSV or XLSX file
 
For PIM-enabled groups, the tool can also configure the group's **PIM Member policy** and optionally assign users as **eligible members**.
 
The window automatically adjusts to the available screen size. If all options cannot be displayed at once, scroll through the group configuration section. The activity log remains available at the bottom of the window.
 
---
 
## Prerequisites
 
Before using the tool, make sure the following requirements are met.
 
### PowerShell
 
Windows PowerShell 5.1 is required.
 
### Microsoft Graph PowerShell
 
Install the Microsoft Graph Authentication module:
 
```powershell
Install-Module Microsoft.Graph.Authentication -Scope CurrentUser
```
 
### ImportExcel
 
The **ImportExcel** module is only required if you want to create or import XLSX files.
 
```powershell
Install-Module ImportExcel -Scope CurrentUser
```
 
CSV files can be used without installing ImportExcel.
 
### Microsoft Entra Licensing
 
Microsoft Entra ID P2 or EMS E5 licensing is required to use Privileged Identity Management (PIM).
 
### Microsoft Graph Permissions
 
When you connect, the tool requests the Microsoft Graph permissions needed to create groups and configure PIM.
 
The requested permissions are:
 
- `User.Read`
- `User.Read.All`
- `Group.ReadWrite.All`
- `RoleManagementPolicy.ReadWrite.AzureADGroup`
- `PrivilegedEligibilitySchedule.ReadWrite.AzureADGroup`
 
Depending on your organization's Microsoft Entra configuration, administrator consent may be required before these permissions can be used.
 
### Administrative Role
 
The account running the tool must have sufficient Microsoft Entra permissions.
 
Creating role-assignable groups and managing their PIM configuration generally requires the **Privileged Role Administrator** role.
 
---
 
## Starting PIM Group Manager
 
Open Windows PowerShell and change to the folder containing `PIM-Group-Manager.ps1`.
 
Launch the tool using:
 
```powershell
powershell.exe -STA -NoProfile -File ".\PIM-Group-Manager.ps1"
```
 
The `-STA` option is required for the graphical interface.
 
---
 
## Connect to Microsoft Entra
 
Before creating or modifying groups, connect the application to Microsoft Entra.
 
1. Review the **Audit log** location shown at the bottom of the window.
2. If you want to store the log somewhere else, select **Browse** and choose another folder and `.log` file name.
3. Select **Connect / Login**.
4. Sign in using the account that will create or manage the groups.
5. If prompted, grant or request administrator consent for the required Microsoft Graph permissions.
6. Select **Check connection**.
 
The connection check verifies:
 
- The signed-in account
- The Microsoft Entra tenant
- The required Microsoft Graph permissions
 
If required permissions are missing, group creation and PIM changes will remain disabled.
 
When you are finished using the tool, select **Disconnect**.
 
---
 
## Audit Log
 
PIM Group Manager keeps a local audit log of activity performed through the application.
 
Each session creates a unique log file in the user's **Documents** folder using a name similar to:
 
```text
PIM-Group-Manager-<timestamp>-<session>.log
```
 
The log file is created when the application starts.
 
### Changing the Log Location
 
Select **Browse** to choose another folder or `.log` file.
 
If you select an existing log file, the application will ask for confirmation before adding new entries. Existing log files are never intentionally overwritten.
 
If the log location is changed during a session:
 
- New activity is written to the new log file.
- Previous activity remains in the original log file.
- Each entry contains a session ID that can be used to identify events from the same session.
 
### Information Recorded
 
Each log entry is stored as a JSON record and can include:
 
- UTC timestamp
- Session ID
- Computer name
- Windows user
- Microsoft Graph account
- Microsoft Entra tenant
- Group creation activity
- Created group IDs
- Group configuration
- PIM policy settings
- Eligible-member requests
- Results
- Errors
 
Passwords and Microsoft Graph access tokens are not intentionally written to the log.
 
> **Important:** Logs can contain information such as user UPNs, group descriptions, and dynamic membership rules. Store logs in an access-controlled location and follow your organization's retention requirements.
 
The local audit log is intended for troubleshooting and operational history. It is not tamper-proof. Microsoft Entra audit logs should be used when authoritative service-side auditing is required.
 
### LOG ERROR
 
If **LOG ERROR** appears, the application was unable to write to the configured audit log.
 
Select a different `.log` file in a location where you have write access.
 
The application blocks new Microsoft Graph changes when required audit information cannot be recorded.
 
However, it is possible for a Microsoft Graph operation to complete immediately before a later log write fails.
 
Before retrying an operation after a logging error, check the group in Microsoft Entra to make sure the change was not already completed.
 
---
 
## Create a Single Group
 
### 1. Enter the Group Information
 
Enter the:
 
- Group name
- Description
- Mail nickname
 
### 2. Select the Group Type
 
Choose one of the following:
 
- **PIM-enabled security (role-assignable)** - Default
- **Security group (assigned membership)**
- **Security group (dynamic membership)**
- **Microsoft 365 group (assigned membership)**
- **Microsoft 365 group (dynamic membership)**
 
### 3. Configure Dynamic Membership
 
This step only applies when creating a dynamic group.
 
Select whether the dynamic membership rule applies to:
 
- Users
- Devices
 
Then enter the Microsoft Entra dynamic membership rule.
 
### 4. Review PIM Settings
 
This section applies only to **PIM-enabled security (role-assignable)** groups.
 
The default PIM Member settings are:
 
- Maximum activation time: **8 hours**
- MFA required when a user activates membership
- Eligible assignments expire after **365 days**
- Permanent eligible assignments are not allowed by default
- Permanent active assignments are not allowed by default
- MFA is required for active assignments
- Justification is required for active assignments
 
If **Active Assignment Days** is left blank, the tool keeps the maximum duration currently configured in the PIM policy.
 
### 5. Add Eligible Members - Optional
 
You can optionally specify users who should be eligible to activate membership in the group through PIM.
 
Enter each user's UPN and separate multiple users with semicolons.
 
Example:
 
```text
user1@contoso.com;user2@contoso.com;user3@contoso.com
```
 
These users are added as **PIM-eligible members**.
 
They are **not** immediately added as active members of the group.
 
### 6. Create the Group
 
Select **Create group**.
 
Review the confirmation information and approve the operation.
 
For PIM-enabled groups, the application automatically adds the following warning to the beginning of the group description:
 
```text
*DO NOT ADD TO THIS GROUP* Add from PIM Assignments portal
```
 
### Important: Creating the Group Does Not Assign an Entra Role
 
Creating a role-assignable group allows Microsoft Entra roles to be assigned to the group.
 
It does **not** automatically assign any Microsoft Entra roles.
 
For example, if the group is intended to provide **Global Reader** access, the Global Reader role must still be assigned to the group separately through Microsoft Entra/PIM.
 
---
 
## Apply PIM Settings to an Existing Group
 
Use this option when the group was successfully created but the PIM configuration did not complete.
 
**Do not create the group again.**
 
Instead:
 
1. Connect to the same Microsoft Entra tenant.
2. Select **PIM-enabled security (role-assignable)**.
3. Enter the existing group's **Object ID** in **Existing group ID**.
4. Review the PIM Member settings.
5. Select **Apply PIM to existing group**.
6. Carefully review the retrieved group name and Object ID.
7. Confirm the operation.
 
After creating a single group, the application automatically places the new group's Object ID into the **Existing group ID** field.
 
For groups created through bulk creation, the Object ID is included in the results file.
 
### Existing Group Requirements
 
The group must be:
 
- Cloud-only
- A security group
- Role-assignable
- Configured with assigned membership
 
The tool will not apply this PIM configuration to an incompatible group.
 
### What This Option Does
 
**Apply PIM to existing group**:
 
- Retrieves the existing group
- Verifies that the group is compatible
- Applies the selected PIM Member policy
- Reads the configured rules back from Microsoft Graph for verification
 
It does **not**:
 
- Create another group
- Assign Microsoft Entra roles to the group
- Add eligible members
 
If a previous attempt failed partway through, check the group and its existing PIM assignments before retrying. Some changes may already have completed successfully.
 
---
 
## Understanding Partial PIM Configuration
 
Creating a PIM group involves several separate Microsoft Graph operations.
 
For example:
 
1. Create the group.
2. Retrieve its PIM policy.
3. Configure the PIM Member policy.
4. Verify the policy.
5. Create eligible-member assignments, if requested.
 
Because these are separate operations, the group can be created successfully even if a later PIM configuration step fails.
 
This does **not** necessarily mean the entire operation failed.
 
Review the activity log and the group in Microsoft Entra before attempting the operation again.
 
For certain Microsoft Graph `HTTP 404` responses or temporarily unavailable policy assignments, the tool automatically retries the operation for a limited period and refreshes the policy ID between attempts.
 
If the retry still fails, the application displays information such as:
 
- Microsoft Graph operation
- HTTP status
- Graph error details
 
Use this information for troubleshooting.
 
You can also finish the PIM configuration manually through the Microsoft Entra PIM portal if necessary.
 
---
 
## Bulk Group Creation
 
Bulk creation allows multiple groups to be created from a CSV or XLSX file.
 
### 1. Create a Template
 
Select:
 
**Create CSV / XLSX template**
 
Choose a location and file type.
 
- **CSV** does not require ImportExcel.
- **XLSX** requires the ImportExcel PowerShell module.
 
### 2. Add the Groups
 
Enter one group per row.
 
`GroupName` is required.
 
Supported `GroupType` values are:
 
```text
PIMRoleAssignable
SecurityAssigned
SecurityDynamic
Microsoft365Assigned
Microsoft365Dynamic
```
 
### 3. Configure Dynamic Groups
 
For dynamic groups, complete:
 
`MembershipType`
 
Supported values:
 
```text
User
Device
```
 
Also enter the required dynamic rule in:
 
`MembershipRule`
 
### 4. Configure PIM Groups
 
Each PIM group can have its own PIM settings in the import file.
 
Boolean fields use:
 
```text
True
False
```
 
Eligible users can be entered in `EligibleMemberUPNs` and separated with semicolons.
 
Example:
 
```text
user1@contoso.com;user2@contoso.com
```
 
Leave `ActiveAssignmentDays` blank if you want to keep the maximum duration currently configured in the PIM policy.
 
### 5. Import the File
 
Select:
 
**Import CSV / XLSX**
 
The application validates the entire file before bulk group creation is enabled.
 
If validation fails:
 
- The previous preview is cleared.
- Bulk creation remains disabled.
- Correct the reported problems in the CSV or XLSX file.
- Import the corrected file again.
 
### 6. Create the Groups
 
After the file successfully passes validation:
 
1. Select **Create all validated groups**.
2. Review the number of groups that will be created.
3. Confirm the operation.
4. Allow the tool to process the groups.
5. Review the results CSV created beside the imported file.
 
The results file contains information about each attempted group creation, including the created group ID when available.
 
### XLSX Template Placeholder
 
The XLSX template contains a placeholder group named:
 
```text
__REPLACE_ME__
```
 
Replace this value with an actual group name or remove the row before importing the file.
 
Completely blank template rows are ignored.
 
---
 
## Important Limitations and Considerations
 
### Role-Assignable Groups Cannot Use Dynamic Membership
 
Microsoft Entra role-assignable groups must be:
 
- Security groups
- Configured with assigned membership
 
They cannot use dynamic membership.
 
If dynamic membership is required, create a separate non-role-assignable dynamic group.
 
### Dynamic Membership Rules Are Not Built by the Tool
 
The application submits the dynamic membership rule to Microsoft Entra as entered.
 
Always validate the rule and confirm that it will select the intended users or devices before creating the group.
 
### Group Creation and PIM Configuration Are Separate Operations
 
The following actions are separate Microsoft Graph operations:
 
- Group creation
- PIM policy configuration
- Eligible-member assignment
 
A failure during a later step does not automatically remove a group that was already created.
 
The activity log and bulk results identify partial completion when possible.
 
### Verify PIM Configuration After Creation
 
PIM Group Manager configures PIM through the Microsoft Graph role-management policy APIs.
 
After creating or modifying a PIM group, verify the resulting configuration in the **Microsoft Entra PIM portal**.
 
This is especially important when:
 
- A Microsoft Graph request returned an error
- Automatic retries occurred
- Eligible-member creation partially failed
- PIM settings were reapplied to an existing group
 
---
 
## Recommended Workflow
 
For a new PIM-enabled group:
 
1. Start **PIM Group Manager**.
2. Review the audit log location.
3. Select **Connect / Login**.
4. Select **Check connection**.
5. Enter the new group information.
6. Leave **PIM-enabled security (role-assignable)** selected.
7. Review the PIM Member settings.
8. Optionally enter users who should receive eligible membership.
9. Select **Create group**.
10. Review the activity log for any warnings or errors.
11. Verify the group and PIM configuration in Microsoft Entra.
12. Assign the required Microsoft Entra role to the group separately.
13. Test PIM activation with an eligible user before using the group for production access. 
