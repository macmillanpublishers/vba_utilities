# Overview
End users can install the template themselves following the instructions [on this page](https://confluence.macmillan.com/display/PBL/Install+the+Macmillan+Template). This document just explains what the installer file does for troubleshooting or manual installation purposes (or just because you're curious).

The installer macro downloads the template files that are attached to  [this Public Confluence page](https://confluence.macmillan.com/display/PBL/Word+Template+downloads+-+production) and saves them in the correct locations on the user's computer. It is a sibling page to the Public space home page, so users won't see it in the page tree when viewing the home page.

Note that the VBA code needs to reference the attachment URL, not the main page URL, i.e. the URL for the main attachments page is <https://confluence.macmillan.com/download/attachments/9044274>. Simply append the file name to the end of the URL to download that file.

# Git Confluence connector
## Production files
The live production files (templates and version number files) are attached to the page listed above using the [Git for Confluence](https://marketplace.atlassian.com/plugins/nl.avisi.confluence.plugins.git-plugin/server/overview) add-on. This keeps the files synced with the `master` branch of the `Word-template` repo on GitHub. Any changes that are pushed to the `master` branch are automatically available for download.

Admin docs for the add-in [are available here](http://addons.avisi.com/git-for-confluence/documentation/).

## Beta testing and development files
There are separate Confluence pages whose attachments are synced to the `releases` branch [(here)](https://confluence.macmillan.com/display/PBL/Word+template+downloads+-+pre-release) and the `develop` branch [(here)](https://confluence.macmillan.com/display/PBL/Word+template+downloads+-+staging). See Development Workflow for their uses.

**Note:** The *downloadBranch* variable in the `AutoExec` procedure in the `ThisDocument` module of each template file (`MacmillanGT.dotm`, `GtUpdater.dotm`, and the installer file) needs to be set to the branch you want to download from. See [Development Workflow](development+workflow) for the process to manage this.


## Scheduled jobs
The Git for Confluence add-on checks the repo for updates every five minutes. A Confluence administrator can sync the files manually by going to **General configuration > Administration > Scheduled Jobs**. One of the jobs in the list is **Synchronize Git repositories**, and it lists the time of the last sync and the time of the next scheduled sync. To sync the files manually, click on the **Run** link.


# Installer file
Users download the installer file from the link on [this Confluence page](https://confluence.macmillan.com/display/PBL/Install+the+Macmillan+Template), which downloads the file directly from the `master` branch of the GitHub repo.

When you open the installer file, the `Document_Open` procedure in the `ThisDocument` module is executed, though with default security settings most users will be prompted with a warning and need to click **Enable Content** before it will run.

If you need to open the file without running the macro (to edit code, for example), hold down Shift while you open it.

# Update checker
The templates check once a day to see if new versions are available. They do this by downloading the version files attached to the appropriate Confluence page and comparing the number listed to the number in the *Version* custom document property in each template. If the template version is less than the posted version, the user is prompted that an update is available. If they click OK, the macro then downloads the new version.

If the file that is being checked is missing at any time (not just during the once-a-day check), the user will also be prompted to download the file again.


# Installed file locations
There are a number of template files that need to be saved in specific directories for the macros to work correctly.

## Startup
Word automatically loads any templates saved in the Startup directory  as Add-Ins when it launches. As soon as it is loaded, it will execute any `AutoExec` procedures in the `ThisDocument`.

### Files to be installed here
* **GtUpdater.dotm:** Has an `AutoExec` procedure that loads `MacmillanGT.dotm` as an Add-In and also checks daily that `MacmillanGT.dotm` is up to date, then prompts the user and downloads the updated files if needed.

### Location
Note that because the Startup directory can be changed by the user, the paths listed below are just the defaults. The installer macro actually uses the `Application.StartupPath` property to locate it.

#### PC
* `C:\Users\<username>\AppData\Roaming\Microsoft\Word\STARTUP` (where `<username>` is replaced by the actual username)

You can change the current Startup directory in Word 2010 and 2013 via **File > Options > Advanced > General > File Locations... > Startup**.

#### Mac
* `Macintosh HD:Applications:Microsoft Office 2011:Office:Startup:Word`

You can change the current Startup directory in Word 2011 via **Preferences > File locations > Startup**.

Note that the Startup directory is read-only for Mac users without admin access, which is why Mac users in-house install the template via Self Service.


## MacmillanStyleTemplate
All other templates must be stored in the `MacmillanStyleTemplates` directory. If it doesn't already exist, the installer or updater macro will create it.

Note that the directory name must be spelled exactly as listed (camel case, no spaces) for the macros to work correctly.


### Location
#### PC
The MacmillanStyleTemplate directory is located in `C:\ProgramData\` (the root being of course the user's local drive). Note that `ProgramData` is a hidden directory. 

#### Mac
The MacmillanStyleTemplate directory is located in the user's Documents directory (e.g., `Macintosh HD:Users:<username>:Documents`). 


### Files to be installed here
* **MacmillanGT.dotm:** the primary template, containing the Macmillan custom macros and custom ribbon tab.
* **macmillan.dotm:** the primary style set, containing over 400 custom manuscript styles.
* **macmillan_NoColor.dotm:** exactly the same styles as `macmillan.dotm` but with colored borders and shading removed.
* **macmillan_CoverCopy.dotm:** styles for use in jacket / cover copy documents.

# Other files you may see
These files may appear in the same directories as the template files, but do not need to be available for the template to work.

## Version files
Each template has a text file of the same name, which contains the current version number of that template. When the templates check for updates each day they save that file in `MacmillanStyleTemplates`.

## Log files
A subdirectory named `log` is created by the installer macro in the `MacmillanStyleTemplate` directory. It contains logs of downloads for each template file.

The Castoff macro also saves CSV files of design info here, and the Bookmaker check macro saves a CSV of current Bookmaker styles.









