Word comes with a built-in IDE for working with macros, the Visual Basic Editor (VBE). It's not perfect, but it gets the job done. This is some advice for working on VBA code in general, and using our workflow specifically.

This page assumes you have basic familiarity with the VBE; if not, you can check out [Microsoft's documentation](https://msdn.microsoft.com/en-us/library/office/jj692815.aspx).

# Opening / saving template files
## Open templates with File > Open method
If you double click on a *template* file, it creates a new *document* based on the template; to edit the template directly, you need to use **File > Open**.

## Don't save template files on a Mac
When you save a template file on a Mac, any custom UI elements (i.e., the Macmillan Tools tab in `MacmillanGT.dotm`) are removed. We use the following workaround to avoid saving a template on a Mac:

Do primary development for all templates on a PC. When you need to make changes on a Mac, do the following:
1. Export *only* the code module(s) to your local repo--do not save the updated template file. 
1. Commit and push the changes.
1. Pull the changes to the local repo on your PC.
1. Run the Import All Modules utility macro on the PC to import the updated modules and save the updated templates in the local repo.
1. Commit and push the changes.
1. Pull the updated template files down to the local repo on your Mac. 

## Always work on file in Startup or MacmillanStyleTemplate
Do not work directly on the template files that are in your local Git repo; instead, work on the template files in their final installation locations. This addresses multiple concerns, including: (1) always working on the same file helps avoid confusion as to which file is the most up-to-date; (2) many macros rely on files being in specific locations, and working from the template in its final location makes it clearer that we're referencing the correct files; (3) the templates are still loaded as Add-Ins even if you open a copy from a different location, which causes duplicate versions of the file to be open.

Both the Import All Modules and Export All Modules utility macros copy the template files from the open location to the Git repo, so maintaining an up-to-date template version in both locations is quite simple.


# Dev utilities
As explained [elsewhere](code+organization), we use Git for version control, and to do this we need to export all code modules to the local Git repo to track and merge changes. Also, some templates share modules, so we have a [very specific workflow](development+workflow) to keep thing everything up to date. 

To better manage this process, we have a couple of helpful macros stored in the `Utilities.dotm` template in the `Word-template` repo. 

## Install Utilities.dotm
1. Quit Word.
1. Copy the file `Utilities.dotm` from the local repo to your Word Startup directory.
1. Launch Word.
1. Use **File > Open** to open `Utilities.dotm` from the *Startup* directory (*not* from the Git repo directly).
1. Open the VBE and view the `Utilities` code module.
1. Edit the location of your local `Word-template` repo for the constant *strRepoPath* at the top of the module.
1. Save the template file.

Now you should have a few buttons added to the Quick Access Toolbar, which launch the macros described below.

## Troubleshooting
If these aren't working as expected, it might be helpful to [check out this page](http://www.cpearson.com/excel/vbe.aspx) on programming the VBE. Also, keep in mind:

* If you get an error when you try to run the macros, go to **File > Options > Trust Center > Macro Settings**  and check "Trust access to the VBA project object model."
* This hasn't been tested on a Mac, just on PC, so use at your own risk.
* Because VBA used to deploy viruses often manipulates the VBE, some anti-virus programs will automatically delete modules that contain code which tries to do just that.

## Macros
### Export All Modules macro
Exports all of the modules in all open templates to the local Git repo. Modules that start with *Shared* in their name get saved to `Word-template/SharedModules`; the rest get saved to the same directory as the template. The template itself is also copied to the repo.

You can run this by clicking the blue up arrow in the Quick Access Toolbar.

#### Shared Modules
If you've changed one of the shared modules, you should run this macro with only the template containing the changes open. Otherwise it may overwrite the changed module with a previous version when it exports the module from a different template.

It's then a good idea to open the other templates and run the Import All Modules macro (below) to make sure all templates are up to date.

#### Userforms
Userforms export as two files (.frm and .frx). Sometimes the .frm file will include an additional blank line 

#### Export modules manually
To export an individual module without using this macro:

1. In the Projects Window, right-click or Cmd-click on the module you'd like to export.
1. From the list that opens, select **Export File...**
1. Navigate to the correct subdirectory of your local Git repo.
1. Click **Save**.
1. If prompted that the file already exists, choose to replace the old version of the file.

### ImportAllModules macro
Removes clears all the code in all open templates, and then imports (1) all of the modules in each template's repo subdirectory, and (2) all of the modules in `SharedModules`. 

This is particularly useful for [merge conflicts](assorted+best+practices) with template files and keeping all templates shared modules up to date.

You can run this by clicking the blue down arrow in the Quick Access Toolbar.

#### Import modules manually
To import an individual module without using this macro:

1. In the Project Explorer, right-click or Cmd-click on the module you'd like to import (if it already exists).
1. Select **Remove _ModuleName_** from the menu (where _ModuleName_ is the name of the module  you're going to import).
1. When prompted "Do you want to export _ModuleName_ before removing it?" select **No**.
1. Back in the Project Explorer, right-click or Cmd-click on the template project name.
1. From the list that opens, select **Import File...**
1. Navigate to the correct subdirectory of your local Git repo and select the module you want to import.
1. Click **Open**.

Note the `ThisDocument` class module is a permanent feature in each template and cannot be deleted or have its name changed. Thus, if you need to update one of these modules from an external class file, it is easiest to open the `TheDocument.cls` file in a text editor and then copy and paste the code into the `ThisDocument` module. If you try to import it using the instructions above you'll get a duplicate module.

