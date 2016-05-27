# Branches
Our branching model is [based on this one here](http://nvie.com/posts/a-successful-git-branching-model/), with the tweak that our `releases` branch has an infinite lifetime.

The *downloadBranch* variable in the `AutoExec` procedure in the `ThisDocument` module of each template file (`MacmillanGT.dotm`, `GtUpdater.dotm`, and the installer file) needs to be set to the name of the branch to download updates from the correct place.

### master
Current production files. The template and version files that are attached to [the production download page](https://confluence.macmillan.com/display/PBL/Word+Template+downloads+-+production) (i.e., from which the installer downloads when installing for end-users) are synced directly to the repository with the Git for Confluence connector.

Only `hotfix-*` branches may branch off of `master`, and only `hotfix-*` and `releases` branches may be merged into it. By definition, any changes merged into `master` are a new release and require that the version number be incremented.

### releases
Used for beta testing. These files are attached to [this page](https://confluence.macmillan.com/display/PBL/Word+template+downloads+-+pre-release) page via the Git for Confluence connector. 

Primarily a way for our composition vendor (and any brave and interested individuals) to test updates before they get a wider release.

Nothing should branch off of this branch, and only `hotfix-*` and `develop` may be merged into it. 

### develop
Development branch. As much as possible, should be stable enough to be ready for the next release.  These files are attached to [this page](https://confluence.macmillan.com/display/PBL/Word+template+downloads+-+staging) page via the Git for Confluence connector.

Only `features` branches may branch off of it, and only `features` and `hotfix-*` branches may be merged into it.

### features
Individual features are developed in their own branches with whatever name makes sense. The branches are deleted after the feature is merged into `develop`.

### hotfix-*
Created when quick changes need to be made to `master` for urgent fixes that do not incorporate the current `develop` branch. The * in the branch name should be replaced by the new version number. Hotfix branches are deleted after they have been merged back into `master`.



# Detailed steps for merging branches
The specific steps involved in the development process are a little more complicated than presented [here](http://nvie.com/posts/a-successful-git-branching-model/), because (1) we need to manage the template binary files, and (2) some modules are shared among templates.

## New hotfix changes
* Begin a new `hotfix-*` branch from `master`, using the new version number in the branch name. 

```
$ git checkout -b hotfix-2.0.1 master
```

* Copy the template files from the Git repo to their final install locations.
* Make whatever changes are required.
* When you're happy with your changes, test *all* macros on both PC and Mac to make sure they are working.
	* If you make any changes to any of the shared modules, open just the template that contains those changes and use the Export All Modules macro.
	* Open all of the templates (don't forget the installer file!) and use the Import All Templates macro to update the shared modules in all templates.
* Open all code templates (don't forget the installer file!).
* Use the Import All Modules macro.
* Open all code templates again.
* Verify that the *downloadBranch* variable in the `AutoExec` and `Document_Open` procedures is set to `master` in each template.
* Use the Export All Modules macro.
* Use the version update macro to update the version number(s).
* Commit and push the changes.
* Open a pull request.
* When the pull request is closed, wait at least 5 minutes for the files to sync to the download page, and then download the installer file from Confluence and run it to test that everything is working.
* Checkout the `releases` branch and merge the `hotfix-*` changes into it.
* Checkout the `develop` branch and merge the `hotfix-*` changes into it. You may need to fix merge conflicts (see below).
* Delete `hotfix-*`.

```
$ git branch -d hotfit-2.0.1
```

## New features
* Begin a new `feature` branch from `develop`, using whatever name you want.

``` 
$ git checkout -b newfeature develop
```

* Copy the template files from the local repo to the final install locations.
* Change the variable *downloadBranch* in the `AutoExec` and `Document_Open` procedures in each template to `develop`.
* Work on your new feature, exporting modules and pushing changes as needed.
* When the new feature is stable and you want to add it to the next release, merge it into the `develop` branch.

```
$ git checkout develop
$ git merge --no-ff newfeature
```

* You may get merge conflicts here. Follow the instructions below to resolve them.
* Delete the feature branch.

```
$ git branch -d newfeature
```

* Test all of your macros on PC and Mac to make sure they are still working.

## New release
* Make sure your `develop` branch is stable.
* Checkout the `releases` branch and merge `master` into it, to verify that there won't be conflicts when the release goes live.

```
$ git checkout releases
$ git merge --no-ff master
```

* Merge conflicts here should be rare, but fix them if they occur.
* Merge `develop` into releases.

```
$ git merge --no-ff develop
```

* You will likely get merge conflicts here. Resolve them following the instructions listed below.
* Change the variable *downloadBranch* in the `AutoExec` and `Document_Open` procedures in each template to `releases`.
* Export All Modules, and commit and push the changes.
* Test all of your macros on PC and Mac to make sure they are still working.
* Update the version numbers for changed templates.
* Commit and push the changes.
* Follow notification instructions below. The new version number will trigger updates for anyone on the `releases` branch.
* Incorporate any fixes revealed from beta testers.
* For each fix, increment the version number so they get the updates to continue testing.
* Once the `releases` branch is stable, open all templates and Export All Modules.
* Commit and push the changes, if any.
* Open a pull request.
* When the pull request is closed, wait at least 5 minutes for the files to sync to the download page, and then download the installer file from Confluence and run it to test that everything is working.
* If any changes were made during beta testing, checkout the `develop` branch and merge the `releases` changes into it.

```
$ git checkout develop
$ git merge --no-ff releases
```


# Merge conflicts
Git cannot merge changes to binary files, so if you have merged different branches you may get a merge conflict for your Word template files. Luckily you exported your code modules, which *can* merge two different branches, so just follow the steps below to update the template files.

```
$ git checkout develop # or whatever branch you are merging changes into
$ git merge --no-ff feature # "feature" being whatever feature branch you are merging in
```

You may get merge conflicts here. Resolve any conflicts in the code files (.bas, .cls, or .frm) by editing them in a text editor, then committing the changes to those files.

Next, we will resolve the binary file conflicts (document templates or .frx userform files) by selecting the file from only one branch using the `--ours` or `--theirs` option with `git checkout`. If all that changed was code, it doesn't matter which version you choose, but it does matter for things like styles, userforms, version numbers, and the like.

```
git checkout --theirs path/file.dotm
```

Now, open the template files and use the Import All Modules macro to import the merged code modules back into the template files. If you made any changes to the custom Ribbon tab, you'll want to re-import that XML code as well.

Any time you do this, be sure to double check that (1) all macros work correctly (i.e., no errors were introduced when you fixed the merge conflicts in the text editor), and (2) the version numbers are correct.

Then you can save the template, and commit the changes.


# Deployment
Updates to `master` will be available within five minutes of closing the pull request. However, the auto-update macro only checks for updates once a day, so most users won't be prompted to update the template until the following day.

## Notifications
Before submitting a pull request to merge `releases` into `master`, the following groups must be notified.

### Westchester (composition vendor)
email:  APS@antares.co.in, Macmillan_PreEdit@wbrt.com, Tina.Mingolello@wbrt.com, Terry.Colosimo@wbrt.com

* 24-48 hours before release for bug fixes
* 5-7 business days before release for template and/or macro updates, especially if they involve process changes 

### Users
Send notification via MailChimp to local users (using the "Word Styles Newsletter" segment) with basic info about the updates (24-48 hours before deploying new release).

### Mac Support
If changes have been made to `GtUpdater.dotm`, a ticket should be opened with Mac Support to update this file in Self Service.
