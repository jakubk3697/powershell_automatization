# Active Directory Automatization in PowerShell

### Little console program which can be used for faster creation organizational units, groups and user account in this places. Also you can generate specific reports about Active Directory. 
---
**:scroll:USAGE**

**if you perform an action then the logs will be saved in files. All files are created in path `C:\PS\generated`. You can change the path by modifying `$dirPath` variable in script.**
1. Open script in PowerShell on your domain controler with `Windows Server`.
2. Select option to perform the action or type enter with random/empty value to exit submenu.
3. If the script is successful you will see a green prompt:green_heart:, if there is an error you will see a red prompt :exclamation:
---
**:round_pushpin:INFO**
- ***:one: Handle user accounts*** - Redirects to the submenu with scripts for users in AD 
	- **:one: Create a user account** - Takes params `name` `surName` `deparment`and create user in Active Directory.
	-	**:two: Create multiple accounts from a csv file** - In this step you have to write `path` for `.csv` file which contains users with headings in format **`name|surname|department`** or you can use file`users_to_add.csv `  which was created automatically.
	-	**:three: Disable a user account** - Takes param `login` (e.g. `john.doe`) and disable account with this. credential1
	-	**:four: Change user account password** - Takes param `login` (e.g. `john.doe`) and, if the account exists in Active Directory you can write 	new password
-	**:one: Handle group accounts** - redirects to the submenu with scripts for users in AD
	-	**:one: Create a new group** - Takes param `groupName` and creates this group in Active Directory
	-	**:two: Add a new user to the group** - first takes the parameter group name and then the member you want to add
-	**:one:Generate reports** - redirects to the submenu with scripts to generate reports with info from AD
	-	**:one: Generate group lists with members** - for each group in AD, it creates a separate file in the `reportGroupAccounts` folder that contains its members. 
	-	**:two: Generate a list of disabled accounts in the domain** -  generates a report in a file  `disabled_accounts.csv`in `reports` folder including `Account name | DistinguishedName | SID |Date of last modification`
	-	**:three: Generate lists of detailed information about user accounts** - generates a report in a file  `disabled_accounts.csv` in folder `reports` including detailed information about user accounts 
	-	`name|surName|login(UPN)|samacount|localization in ADDS (DN)|creation date|last modification|last login|last password change"`
	-	**:four: Generate lists of detailed information about computer accounts in the domain** - generates a report in a file `xyz.csv` including information about computer accounts in the domain `Computer name|SID|DistinguishedName|Account status|Last password change|Creation date`
	-	**:five: Generate list of organizational units in the domain (info sorted alphabetically relative to the OU)** - generates report including Organizational Units in AD, sorted alphabetically by OU. Report contain info `OU|DistinguishedName`
