The ArcMap Mapbook Project (shortname 'arcmapbook') is an open source and 
volunteer effort to extend, enhance, and fix the bugs in the ESRI developer 
sample, DS Mapbook.

ArcMapbook installation & removal should be as simple as running the _INSTALL 
and _UNINSTALL batch files. For upgrades it is a good idea to unsintall the 
old version first; existing map series will be unaffected.

Visual_Basic: 	The program
Docs: 			How to install and use Mapbook.

Project homepage: http://arcmapbook.sourceforge.net/

-------------------------------------------------------------------------------
Mapbook for ArcMap 9.2, Release 2007 October 26

The main reason for this release is to bundle the slightly improved un/install with the main code. Changes from March release are:

- improved installation and removal usability (moved un/install.bat to top folder)

- initial check-in of Jerry Chase's customisation of Mapbook (Customized DS Map Book Manual.doc). NOTE: actual code is not yet implemented in arcmapbook, you'll have to pull it out of the doc yourself. Jerry's doc includes instruction and code. Outline of added functionality: 1. The ability to control the extent of the locator frames: a. Scaled Indicator: scale by the percentage of area of the record in focus b. Scaled Local Indicator: scaled by the percentage of the area which includes the record in focus and all records contiguous to it c. Index 2 Indicator: the ability to control extent of indicator frame by the area of a secondary index which is tied to the index of the record in focus (State in which a county is located). 

- Added to pause to (un)install batch files so any error messages can be seen. _UNINSTALL.bat now deletes Mapbook registry keys


-------------------------------------------------------------------------------
Mapbook for ArcMap 9.2, Release 9.2 2007 March 13

This release is a wholesale merge of Larry Young's upstream release of 
2007-March-13[*] with our local documents. From Larry's comments, there are a lot 
of fixes to the Export routines.

Almost every file had dozens of changes, though >90% are just line prefixes (e.g. 
73:... to 86:...). so I probably missed things.

If you are currently using DS Mapbook from the 9.2 developer sample kit, you 
should upgrade to this release. 

If you've already installed the March-13 version from the ESRI support forum, 
ignore the 'Visual_Basic' folder, the only value added stuff this package is the 
documentation.

[*]	http://forums.esri.com/Thread.asp?c=93&f=989&t=211395#653249
		http://forums.esri.com/Attachments/23830.zip

--
Matt Wilkie, 2007 March 30