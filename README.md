
<p align="center">
  <img src="https://github.com/AndyGlx/Images/blob/master/Logo%20NEW%20(white%20BG).gif" width="600">
</p>

Software for designing GCode for additive manufacturing (or for other processes that use GCode such as laser cutting)

See www.fullcontrolgcode.com for more information and tutorial videos

Email info@fullcontrolgcode.com for queries, collaboration and support


The software uses Excel as a front end - the code is visible in the Visual Basic application within Excel. No installation required - just download it from this repository (open the [FullControl_GCode_Designer_Heron02d.xlsm](https://github.com/AndyGlx/FullControl-GCode-Designer/blob/master/FullControl_GCODE_Designer_Heron02d.xlsm) page and click the download button).

Since FullControl used VBA macros, you'll need to activate them when you open the file and potentially follow [microsoft's instructions](https://learn.microsoft.com/en-us/deployoffice/security/internet-macros-blocked#remove-mark-of-the-web-from-a-file) to 'unblock' the file (files downloaded from the internet have added security in recent versions of Windows) 

A journal article describing FullControl is included in this repository for download ('FullControl GCode Designer - Author Version.pdf')

Please cite:
```
Gleadall, A. (2021). FullControl GCode Designer: open-source software for unconstrained 
design in additive manufacturing. Additive Manufacturing, 46, 102109.
```

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
