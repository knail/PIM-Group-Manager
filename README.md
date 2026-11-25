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

The **Assignments** tab groups **Query one group**, **Add eligible members**, and **Remove members** under **One group**. **All groups (read-only)** contains **Query all eligible** and **Scan permanent active**. Queries are read-only; adding or removing members changes access.
 
The window automatically adjusts to the available screen size. If all options cannot be displayed at once, scroll through the group configuration section. The activity log remains available at the bottom of the window.

Larger, bold tabs highlight the selected tab in blue. Assignment results resize with the window; **Export results** remains in a separate footer, and the group-action sections scroll when needed.

## GUI examples

### Create one group

Configure group properties, PIM policy settings, and optional eligible members.

![Create one group tab with Eligible Assignments and Active Assignments settings](screenshots/create-one-group.png?v=9433ee6df16a)

### Assignments

Query a group, add or remove members, scan all groups, and export report results.

![Assignments tab with single-group actions, all-group reports, and export results](screenshots/assignments.png?v=078942d38a9d)
 
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
 
The **ImportExcel** module is required to create, import, or export XLSX files.
 
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
- `PrivilegedAssignmentSchedule.Read.AzureADGroup` (permanent-active scan)

Member removal requests additional permissions only when needed: `PrivilegedAssignmentSchedule.ReadWrite.AzureADGroup` for PIM active schedules, and `RoleManagement.ReadWrite.Directory` for direct membership removal from role-assignable groups. The tool asks before reconnecting for these permissions.
 
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

The GUI comes to the front on launch, then behaves like a normal window; it does not remain always-on-top.

Only one instance can run per user in a Windows session, including launches from another PowerShell process or another copy of the script. A duplicate launch warns in the console and stops before resetting any GUI state. In the same PowerShell host, it also attempts to restore and activate the existing window.

In VS Code, selecting the **entire script** and pressing **F8** is supported; do not run only selected internal sections, which bypass startup checks. Prefer the command above in a separate PowerShell process to keep the VS Code PowerShell session available while the GUI is open. VS Code may queue F8 requests while the GUI occupies its PowerShell session; a queued request that runs after the window closes is a new launch, not a simultaneous duplicate.

Closing is blocked during sign-in and long-running operations. Wait for completion and review the activity log before closing. Normal closure or a startup error releases the single-instance guard and disposes the windows so the tool can be launched again. This does not roll back Graph changes or automatically disconnect Graph.
 
Select **How to use** in the main window header to open the instructions in a separate window. The help window closes with the main application.

The guide groups instructions by task and adds spacing between steps and bullet points to make them easier to scan.

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

## Assignments

Connect to Microsoft Graph and open **Assignments**:

1. Under **One group**, enter an exact group display name or object ID and select **Query one group** to see current eligible member and owner assignments. The group must be a cloud-only, role-assignable security group with assigned membership. Use the object ID if names are duplicated.
2. Under **All groups (read-only)**, select **Query all eligible** to find current eligible assignments across role-assignable groups.
3. Select **Scan permanent active** to find permanent PIM active member assignments and review direct members.
4. Review the table and select **Export results** to save findings. Excel (`.xlsx`) is the default when ImportExcel is installed; otherwise CSV is the default. Choose either format in the save dialog. Excel reports and bulk templates use table style **Light20**. Suggested filenames include the completed report type after `YYYY-MM-DD_HHmm-`: `all-eligible`, `all-permanent`, or `one-group-eligible`, followed by `-PIM-Assignment-Report`. Values that could be interpreted as spreadsheet formulas are prefixed with an apostrophe in both formats.

The all-groups eligible report omits PrincipalId and PrincipalType from both the table and export. Names or UPNs may be blank when Graph does not return them; query the group individually to see the principal ID.

All-group scans include **all role-assignable groups**, whether or not they have been onboarded to PIM; PIM-managed groups that are not role-assignable are excluded. Eligible group principals, if returned by Graph, are listed as principals rather than expanded into their members. Results are a point-in-time view, not a continuous audit, and changes or Graph replication during a scan can affect results. A failed scan clears its partial results and reports the error.

The permanent-active report covers membership, not ownership. **Permanent PIM active** means a current assigned (not activated) PIM membership instance with no end date. Other direct members are labeled **verify permanence**: group membership alone does not establish whether access is permanent. A member with a temporary PIM assignment is labeled separately because a concurrent direct assignment cannot be ruled out. Verify all candidates in Entra before taking action. Indirect or nested memberships are not evaluated. The scan uses the Graph beta group-members endpoint because the v1.0 endpoint can omit service principals.

The permanent scan requires the delegated `PrivilegedAssignmentSchedule.Read.AzureADGroup` permission (or its ReadWrite equivalent) and a suitable Entra role, such as directory-scoped Global Reader or Privileged Role Administrator for role-assignable groups. If the permission is missing, the scan offers to reconnect and request consent; an administrator may need to grant it. Other existing operations do not require this additional scope. Reports and exports contain potentially sensitive directory data; store and share them appropriately.

Run the offline, mocked report tests with `powershell.exe -NoProfile -File .\tests\AssignmentReports.Tests.ps1`.

### Add eligible members

Under **One group** in the main Assignments tab, enter an exact group name or object ID before selecting **Add eligible members**. The popup displays that group as bold, non-editable text; cancel and change the main field to choose another group. Enter one or more user UPNs in the popup separated by commas, semicolons, or new lines. Duplicate group names require the object ID. Cloud-only security and Microsoft 365 groups with assigned membership are supported, including non-role-assignable groups; dynamic and on-premises synchronized groups are not.

The start defaults to **Start now** and the end defaults to 365 days from opening the dialog. Uncheck **Start now** to choose a future start, edit the end, or choose permanent eligibility. The date and time controls use the workstation's local time zone and are converted to UTC. The existing group PIM policy may reject the selected duration or permanent eligibility.

Review the resolved group, users, and schedule before confirming. Submitting a request can onboard a group to PIM if it is not already managed; onboarding cannot be undone. This creates eligible membership, not active group membership, and does not change the group's policy. Requests are submitted one at a time: if one fails, later users are not attempted. The message identifies accepted, failed or uncertain, and unattempted users. Check Entra PIM before retrying and verify provisioning status after submission.

This action uses the existing `PrivilegedEligibilitySchedule.ReadWrite.AzureADGroup` delegated scope and requires an appropriate directory role or group ownership. Microsoft documents [supported roles and permissions](https://learn.microsoft.com/en-us/graph/api/privilegedaccessgroup-post-eligibilityschedulerequests?view=graph-rest-1.0).

### Remove members

Under **One group**, enter the exact group name or ID in the main window before selecting **Remove members**. The popup displays that group as bold, non-editable text; cancel and change the main field to choose another group. Enter one or more user UPNs, then select the categories to remove: PIM active schedules, PIM eligible schedules, and/or the direct group members relationship. None is selected by default. Supported groups and UPN separators are the same as for adding eligible members. Owners and indirect or nested membership are not removed.

Review the resolved group, requested users, selected categories, and matching operations before confirming. PIM schedule results can include future assignments. If multiple schedules of the same kind match one user, or a schedule is pending or cannot be classified, the tool stops without changing access; review it in Entra. The preview is a point-in-time view and can change before submission.

**These categories can overlap:** a PIM active assignment or activation also appears in the group's direct members. Removing either can affect active access; deleting the group relationship is not limited to memberships originally added through the groups API. Removing eligibility may not end an existing activation, and users may retain access through other paths. When several categories are selected, one operation can change another operation's target and cause a later operation to fail.

Changes are submitted sequentially and stop on the first error. The outcome identifies accepted requests, the failed or uncertain operation, and operations not attempted. PIM requests can be pending even when accepted; always verify PIM status and group membership in Entra before retrying or assuming access has been removed. The local audit log records each operation.

PIM active removal requires delegated `PrivilegedAssignmentSchedule.ReadWrite.AzureADGroup`; eligible removal uses the existing `PrivilegedEligibilitySchedule.ReadWrite.AzureADGroup`. Removing the direct group relationship uses `Group.ReadWrite.All`, plus `RoleManagement.ReadWrite.Directory` for role-assignable groups. Appropriate Entra roles are also required; a Privileged Role Administrator is the least-privileged supported role for removing members from role-assignable groups. See Microsoft's documentation for [PIM active requests](https://learn.microsoft.com/en-us/graph/api/privilegedaccessgroup-post-assignmentschedulerequests?view=graph-rest-1.0), [eligible requests](https://learn.microsoft.com/en-us/graph/api/privilegedaccessgroup-post-eligibilityschedulerequests?view=graph-rest-1.0), and [group member removal](https://learn.microsoft.com/en-us/graph/api/group-delete-members?view=graph-rest-1.0).

---
 
## Audit Log
 
PIM Group Manager keeps a local audit log of activity performed through the application.
 
The default log is saved in the user's **Documents** folder with a name in this format:
 
```text
YYYY-MM-DD_HHmm-PIM-Group-Manager.log
```
 
If the default log already exists, new entries are appended. Sessions started during the same minute share the file; each entry includes a session ID. The log file is created when the application starts.
 
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

The policy settings separate **Eligible assignments** from **Active assignments**. The Active section is lightly tinted red, and permanent active assignments are marked **NOT RECOMMENDED** because they allow ongoing access without PIM activation. These are policy settings, not a request to add active members; the optional UPN list below still creates eligible assignments only.
 
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
 
Enter each user's UPN and separate multiple users with commas, semicolons, or new lines, just like the Add/Remove members popups.
 
Example:
 
```text
user1@contoso.com;user2@contoso.com,
user3@contoso.com
```
 
These users are added as **PIM-eligible members**.
 
They are **not** immediately added as active members of the group.
 
### 6. Create the Group
 
Select **Create group**.
 
Review the confirmation information and approve the operation.
 
For PIM-enabled groups, the application waits 10 seconds after group creation for Microsoft Entra directory propagation before applying PIM settings.

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
3. Enter the existing group's exact display name or **Object ID**. If multiple groups have the same display name, use the Object ID.
4. Review the PIM Member settings.
5. Select **Apply PIM to existing group**.
6. Carefully review the retrieved group name and Object ID.
7. Confirm the operation.
 
After creating a single group, the application automatically places the new group's Object ID into the **Existing group name or ID** field.
 
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
