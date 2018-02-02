# ActiveDirectoryTool - remains unamed

## Copy the code into the Powershell ISE and hit the play icon.

**Overview**

 - This started as a suggestion of writing a script to add all the members of one group to another group in a loop and lead to an app.
 - This app allows you to search by groups and users to perform tasks that aren't entirely made easy by Active Directory.
 - Such as:
   - Getting user information
   - Copying a selection of users to another group
   - Copying an entire group of users to another group
   - Locating the machine that locked out a user
   - Displaying all the users in a group
   - Displaying all the groups a user belongs to   
 - Paste the code into Powershell ISE and hit _Play_
 
**Info**

 - Permissions are based on the permissions of the user running the app.
 - A log file is created/updated in C:\Temp\ADTool\Log-DATE.txt
 - The app first checks if you have a connection to Active Directory.
 - Most of the app will function with read only permissions of AD - with the exception of:
   - Adding users to groups, this will require you have permission to do so.
   - Searching for the origin of a locked out user, this will require domain admin permissions, as this involves searching the security event logs of a domain controller.
 - Help can be displayed with the button in the top right.
 - Searching for users and groups does not require the full name. _Domain A_ will still return _Domain Admins_
 - The app supports tabbing around form attributes and hitting enter within search boxes and when buttons are highlighted.
 - When searching for a user, the final character is ignored, this is to support searching for users by surname and username. 
   
**Functions**

 - Search Source Group
   - Enter a search term for a group(s) in the search box above and hit _enter_/_Search Source Group_.
   - This will populate the left hand side list box with relevant results.
   - You can only highlight one of these results at a time.
   
 - Search Destination Group
   - Enter a search term for a group(s) in the search box above and hit _enter_/_Search Destination Group_.
   - This will populate the right hand side list box with relevant results.
   - You can only highlight one of these results at a time.
   
 - Search By Name
   - Enter a search term for a user(s) in the search box above and hit _enter_/_Search By Name_.
   - This will populate the left hand side list box with relevant results.
   - You can highlight multiple users at a time.
   
 - User Info
   - Search for a user within the username search box. 
   - Highlight a user in the list box and select _User Info_ to populate the bottom box with AD Attributes for that user.
   
 - Copy Selected
   - When searching for users using minimal characters you are likely to receive more than one result.
   - This function allows you to highlight several of those users and copy them to a highlighted group in the right side list box.
   - You will be asked to confirm the action.
   
 - Show Users
   - This function allows you to list all the users within a group.
   - Search for a source group so the results are placed in the left list box.
   - Highlight one of the results and select _Show Users_ to re-populate the box with the users within that group.
 
 - Memberships
   - This function allows you to list all the groups that a user belongs to.
   - Search for a user using the search box and highlight one of the users.
   - Select _Memberships_ to populate the right box with groups that the user belongs to.
 
 - Copy All Users
   - Search for both a source and destination group so that box boxes are populated.
   - Highlight one group from each box and select _Copy All Users_
   - This will ask you to confirm before adding all the users contained within the source group to the destination group.
   - No users are removed from the source group in doing so.
 
 - Show Lockout
   - This function assumes want to locate the machine that is locking out a user.
   - Search for a user and highlight their name.
   - Select _Show Lockout_ to populate the bottom verbose box with the results from the search.
   - This requires domain admin permissions.
 
 - help 
   - Select _help_ to receive a summary of these instructions in the bottom verbose box.
 
 **Troubleshooting**
 
 - If the first check fails and you are told that you do not have a connection to Active Directory - please ensure:
   - You are connected to the network
   - You have Active Directory Services installed
   - You have the Powershell Module for AD Services installed and available.
     - To test this: open powershell and try _Import-Module Active Directory_
	 - An error implies you don't have the module installed or it is running incorrectly.
 - If you cannot copy members between groups please make sure you are running the app as a user who has the permission to do this.
 - The above also applies to the scenario where you are unable to display the lockout location
 - If your username seach responds with nothing try removing the last character or two from your search and search through the results.
 
 **TO DO**
 
 - Provide this program as an exe
