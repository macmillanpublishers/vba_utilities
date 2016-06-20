# RELEASE NOTES #
## MacmillanGT.dotm ##
### v. 3.1 ###
* Added Castoff macro
* Created Startup and Cleanup functions to include in each macro, which include removing all content controls, books, fields, and hidden text. 
* Added tagging of unstyled paragraphs as TX or TX1 to Character Styles macro.
* Cleanup converts hyphens in number ranges to en-dashes, but not in URLs or numbers with multiple hyphens (e.g., phone numbers, SSNs, ISBNs).
* Cleanup removes superscript from (c), (r), TM.
* Bookmaker Report verifies file name only includes alphanumerics characters, underline, or hyphen.
* Fixed error causing PC to sometimes prompt user to save MacmillanGT.dotm on program launch.
* Generates template paths in a single sub.
* Cleanup and Character Styles clears all formatting from paragraph characters, per Westchester.
* Added separate module to wait for Progress Bar to update before executing more code.

### v. 3.0.7 ###
#### March 14, 2016 ####
* Fixing error with installer.

### v. 3.0.6 ###
#### March 11, 2016 ####
* Fixed hard-coded template location in Print Styles macro.
* Add check that download directory is not read-only.
* Remove all hidden text in Cleanup and Character Styles macros.
* Add hyphen + space as something to tag w/ *span preserve characters*.


### v. 3.0.5 ###
#### February 5, 2016 ####
* Converting all Mac directories (e.g. Documents, tmp) to AppleScript
* Moving all progress bar code (for Mac and PC) to `ProgressBar.frm`
* adding `ProgBarHelper.bas` single sub to increment progress bar for any OS

### v. 3.0.4 ###
* Adding Macmillan Tools toolbar to Mac programmatically when new MacmillanGT.dotm file is installed.
* Moving MacmillanGT.dotm install location to MacmillanStyleTemplate. 
* Adding "mass market paperback" as allowed format in Bookmaker Report. 
* Tagging all underline style types in Character Style Macro. 
* Making `Declare` statements dependent on Win64 system. 
* Adding check for text in CTNP paragraphs to Bookmaker Report. 
* Stripping all character styles/formatting from CN paragraphs in Character Style Macro.
* Determining location of MacmillanStyleTemplate on Mac with AppleScript. 
* Allowing *Section Break (sbr)* style after *Page Break (pb)*. 
* Sorting styles-in-use list in Reports by page of first use. 

### v. 3.0.3 ###
#### October 29, 2015 ####
* *span strikethrough characters (str)* added to Tag Character Styles macro
* View Styles macro now toggles everything based on if the current view is Print or Draft, rather than toggling each window individually.
* Bookmaker Check downloads list of good styles from Confluence.

### v. 3.0.2 ###
#### October 19, 2015 ####
* You can now prevent the Cleanup Macro from removing soft returns by using the *span preserve characters (pre)* style.
* We fixed an issue that was causing the installer file to fail for some PC users.
* We fixed an issue that was causing the Style Report to incorrectly flag errors with epigraph styles.

### v. 3.0.1 ###
#### September 22, 2015 ####
* Fixed bug causing Cleanup and Character styles macros to error if Endnotes but not Footnotes.

### v. 3.0 ###
#### September 1, 2015 ####
* Introducing the custom Macmillan Tools ribbon tab to launch all macros.
* All macros fixed so they can run without style template attached.
* Due to above, all macro modules moved to MacmillanGT from macmillan.dotm.
* New and improved `FileInstaller.bas` module added to check for updates daily and install GtUpdater and all style template.
* Macros that are shared among different modules moved to `SharedMacros.bas` module.
* Moved Castoff and CIP application tags to this template.
* Added standard Progress Bar to CIP application macro.
* Added ASCII Triceratops easter egg.
* Cleanup, Character Styles, Style Report, and Bookmaker Report all loop through footnotes and endnotes. 
* Added support for built-in *Footnote Text* and *Endnote Text* styles to Reports.
* Style Report and Bookmaker Report include page of first use for paragraph styles.
* Style Report and Bookmaker Report remove *space ISBN (isbn)* style from any non-ISBN text. 
* Style Report and Bookmaker Report now include character styles.
* Adding Print Styles in Margin macro
* Assorted bug fixes.

### v. 2.2.2 ###
#### April 17, 2015 ####
* Separating macros into own modules

### v. 2.2.1 ###
#### April 10, 2015 ####
* Fixing message box that dropped out

### v. 2.2 ###
#### April 8, 2015 ####
* Changing Confluence address in Attach Template macro
* Adding Check MacmillanGT and Check Macmillan macros, which tell the user the version number of their installed templates.

### v. 2.1 ###
#### April 1, 2015 ####
* Updated Mac macro launch buttons to be more stable and user friendly.
* Added third macro to attach the Cover Copy Template.
* Edited Attach Template Macro so that it only updates the template that is being attached. Previously, if you tried to attach macmillan.dotm or macmillan_NoColor.dotm and either was out of date, the macro would download both files. Now it is easier to add more templates to attach. 

### v. 2.0 ###
* The attach template macro now checks for updates to macmillan.dotm and downloads the updated file if it is available. 
	* The user is prompted before the download and may cancel the update if they want. 
	* A log file is created locally. 
	* The macro only checks for updates once a day. 

### v. 1.1 ###
* Added second attach macro called Remove Color Guides for macmillan_NoColor.dotm template, for use removing the color guides from a styled manuscript. 

### v. 1.0 ###
* Contains macro to attach the macmillan.dotm template to the current document. Save in Startup folder. 

- - -

## GtUpdater.dotm ##
#### March 14, 2016 ####
* Fixing error with installer.

### v. 1.0.4 ###
#### March 11, 2016 ####
* Add check that download directory is not read-only.

### v. 1.0.3 ###
#### Feb. 5, 2016 ####
* Converting all Mac directories (e.g., Documents, tmp) to AppleScript

- - -

## macmillan.dotm / macmillan_NoColor.dotm ##
### v. 4.5 ###
#### March 29, 2016 ####
* added *bookmaker tighten (bkt)* and *bookmaker loosen (bkl)* styles.

### v. 4.4.1 ###
#### March 11, 2016 ####
* Changed formatting of *span small caps* to be full caps at a smaller point size, per Issue #135.


### v. 4.4 ###
#### Feb 5, 2016 ####
* adding *Extract Source (exts)* style
* adding *Titlepage Logo (logo)* style

### v. 4.3 ###
#### January 21, 2015 ####
* *Bookmaker Processing Instruction (bpi)* style added

### v. 4.2 ###
#### October 29, 2015 ####
* *span strikethrough characters (str)* added

### v. 4.1.1 ###
#### October 26, 2015 ####
* Replacing en-dashes in epigraph style names with proper hyphen to prevent errors in Reports macros.

### v. 4.1 ###
#### October 19, 2015 ####
* New *span run-in computer type (comp)* character style, for all of your text message / email text design needs. 
	* You only need to use this for designs where the email / text message is run-in to a body paragraph and is supposed to be a different font than the body text. 
	* This complements the current *Text - Computer Type (com)* paragraph style.

### v. 4.0 ###
#### September 22, 2015 ####
* Moved all macros to MacmillanGT version 3.0 for custom ribbon.
* Removed Macmillan endnote and footnote styles because built-in note styles are preferred.
* Added formatting to built-in note styles.
* Added *Bookmaker Page Break (br)* style.
* Added *Space Break - Print Only (po)* style for Bookmaker.

### v. 3.9.3 ###
#### May 27, 2015 ####
* Changing macro check of attached template to check for specific Macmillan style because it was failing for some users.

### v. 3.9.2 ###
#### May 20, 2015 ####
* adding backmatter styles to list of acceptable styles for Bookmaker Requirements Macro

### v. 3.9.1 ###
#### May 11, 2015 ####
* fixing typo in *Chap Title Nonprinting (ctnp)* style name in Reports macro module

### v. 3.9 ###
#### May 6, 2015 ####
* Adding progress bar to Cleanup, Character Styles, Style Report, Bookmaker macros
* Adding check of illustration/caption/source order to Bookmaker macro
* Adding most Bookmaker checks to Style Report macro as well
* View Styles macro now a toggle
* Assorted bug fixes and improvements

### v. 3.8.5 ###
#### April 17, 2015 ####
* Fixing error in Character styles macro

### v. 3.8.4 ###
#### April 16, 2015 ####
* Separating macros into own modules

### v. 3.8.3 ###
#### April 13, 2015 ####
* Adding content control handling to Bookmaker macro

### v. 3.8.2 ###
#### April 10, 2015 ####
* Adding track changes handling to bookmaker reqs, cleanup, and character styles macros

### v. 3.8.1 ###
#### April 7, 2015 ####
* Fixing bug in Character Styles macro that was causing manual page breaks to drop out.

### v. 3.8 ###
#### April 1, 2015 ####
* Updated Mac macro launch buttons to be more stable and user friendly.
* Bookmaker Check Macro added to template.
* *Chap Number Nonprinting (cnp)* style has been removed and new style *Chap Title Nonprinting (ctnp)* has been added in its place.
* New character styles added:
	* *span symbols ital (symi)*
	* *span symbols bold (symb)*
* *Design Note (dn)* may now contain blank paragraph returns
* Minor formatting fixes

### v. 3.7 ###
* Split Cleanup Macro into two macros, added new toolbar button to launch:
	* Macmillan Manuscript Cleanup Macro
		* now removes space between ellipsis and closing double or single quote
		* now replaces double period with single period
		* now replaces double commas with single comma
		* All actions of Char Styles Macro removed from Cleanup Macro (except bookmarks–both macros remove them)
	* Macmillan Character Styles Macro
		* Removes bookmarks
		* Removes all active hyperlinks
		* Styles all URLs with Macmillan hyperlink character style
		* Unstyled page breaks removed
		* Unstyled blank paragraphs removed, including if last paragraph of document
		* Character styles applied

### v. 3.6 ###
* Added styles for Series Page:
	* *Series Page Author (sau)*
	* *Series Page Heading (sh)*
	* *Series Page List of Titles (slt)*
	* *Series Page Subhead 1 (sh1)*
	* *Series Page Subhead 2 (sh2)*
	* *Series Page Subhead 3 (sh3)*
	* *Series Page Text (stx)*
	* *Series Page Text No-Indent (stx1)*

### v. 3.5 ###
* Style Report macro now opens the style report automatically when it is complete. 
* New paragraph style added: *Chap Number Nonprinting (cnp)*
	* For use on manuscripts that have chapter breaks but do not have chapter numbers or chapter titles. 
	* Chapter numbers should be added to the manuscripts and tagged with this style. The added chapter numbers will not appear in the print book. 
* New character style added: *span ISBN (isbn)*
	* To be used to tag the ISBN on the copyright page. Needed for bookmaker project. 
* New character style added: *bookmaker force page break (br)*, for use in bookmaker project.
* New character style added: *bookmaker keep together (kt)*, for use in bookmaker project.
* New paragraph style added: "About Author Text Head
#### atah)" per request from production editorial. ####
* Assorted formatting fixes.

### v. 3.4.4 ###
* Removed em-dashes from all signature and source lines per request from design. 

### v. 3.4.3 ###
* fixed Cleanup macro (StyleHyperlink sub was removing all blank paragraphs; now it's not).

### v. 3.4.2 ###
* reinstated PC macro buttons that fell out when template was saved on a Mac

### v. 3.4.1 ###
* added styles for text with space "around" (i.e. before AND after)
* *span symbols (sym)* set to Arial per request from Westchester

### v. 3.4 ###
* added styles for text with space before/after
* deleted numbered colum styles for Tables, Charts, Figures
* replaced Column styles with *Column Text (coltx)* and *Column Text No-Indent (coltx1)*
* assorted formatting fixes

### v. 3.3.1 ###
* fixed style names in macros

### v. 3.3 ###
* removed special characters in style names

- - -

## MacmillanCoverCopy.dotm ##

### v. 2.0 ###
#### August 25, 2015 ####
* Moved all macros to MacmillanGT version 3.0 for custom ribbon.

### v. 1.2 ###
#### May 6, 2015 ####
* Adding updated macro modules: progress bar, style report

### v. 1.1.2 ###
#### April 17, 2015 ####
* Separating macros into own modules

### v. 1.1.1 ###
#### April 10, 2015 ####
* adding track changes handling to macros

### v. 1.1 ###
#### April 9, 2015 ####
* adding author name and subtitle character styles

### v. 1.0 ###
#### April 1, 2015 ####
* First draft of template for adding styles to cover copy