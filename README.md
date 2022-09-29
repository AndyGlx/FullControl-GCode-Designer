
<p align="center">
  <img src="https://github.com/AndyGlx/Images/blob/master/Logo%20NEW%20(white%20BG).gif" width="600">
</p>

Software for designing GCode for additive manufacturing (or for other processes that use GCode such as laser cutting)

See www.fullcontrolgcode.com for more information and tutorial videos

Email info@fullcontrolgcode.com for queries, collaboration and support


The software uses Excel as a front end - the code is visible in the Visual Basic application within Excel (no installation - just download and open the file FullControl_GCode_Designer_Heron02d.xlsm by clicking its name in the list of files towards the top of this page)

A journal article describing FullControl is included in this repository for download ('FullControl GCode Designer - Author Version.pdf')



Screenshot:

<kbd><img src="https://github.com/AndyGlx/Images/blob/master/Screenshot.png" /></kbd>



Journal paper figure: 

![alt text](https://github.com/AndyGlx/Images/blob/master/Final%20figure.jpg?raw=true)

Youtube Highlight Video (click to play):

[![IMAGE ALT TEXT](https://github.com/AndyGlx/Images/blob/master/Highlight%20Video%20Thumbnail%20-%20video%20cue.jpg)](https://youtu.be/KlxuZ5JnA0k "FullControl GCODE Designer - Highlight Video")

## Contributing

This repo contains custom git hooks and python scripts to enable all VBA modules, classes, forms, and custom RibbonUI xml files to be exported prior to each commit. You will find the exported files in the .vba and .xml directories.

If you wish to submit a pull request please clone the repo, edit the xlsm file, then run the following command, in the repo directory, to enable the custom hooks prior to any commits:

`git config core.hookspath .githooks`

It may be necessary to do the following in Excel to allow git to access the vba scripts... 
Open FullControl .xlsm file > File > Options > Trust Centre > Trust Center Settings > Macro Settings > Check the box for “Trust access to the VBA project object model” > Close Excel

Alternatively, directly edit the VBA/XLM files in the .vba and .xlm folders - then there is no need to use the githooks. 

This process for contribution is still being optimised so feel free to suggest improvements to it. 

Only changes to the text of vba/xml code will be accepted, to ensure changes can be effectively reviewed. 
For security reasons, changes to the main .xlsm file and .frx files will not be accepted without prior communication. 
