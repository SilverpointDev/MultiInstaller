# Silverpoint MultiInstaller

Silverpoint MultiInstaller is a multi component package installer for Embarcadero Delphi and C++Builder.
It was created to ease the components installation on the IDE.

Silverpoint MultiInstaller can help you install multiple component packs in a few clicks.
Just download the zips and select the destination folder, all the components will be uninstalled from the IDE if they were previously installed, unziped, patched, compiled and installed.
It can also install multiple packages directly from GIT repositories.


For more info go to:  
www.silverpointdevelopment.com


## License

The contents of this package are licensed under a disjunctive tri-license giving you the choice of one of the three following sets of free software/open source licensing terms:

 - Mozilla Public License, version 1.1
   http://www.mozilla.org/MPL/MPL-1.1.html
 - GNU General Public License, version 2.0
   http://www.gnu.org/licenses/gpl-2.0.html
 - GNU Lesser General Public License, version 2.1
   http://www.gnu.org/licenses/lgpl-2.1.html

Software is distributed on an "AS IS" basis, WITHOUT WARRANTY OF ANY KIND, either express or implied.

The initial developer of this package is Robert Lee.


## Installation

Requirements:
- RAD Studio XE or newer


## Getting Started

To install a component pack with MultiInstaller you have to follow these steps:
1) Read the licenses of the component packs you want to install.
2) Get the zip files of component packs.
3) Get the Silverpoint MultiInstaller.
4) Get the Setup.ini file for that component pack installation or create one.

For example, if you want to install TB2K + SpTBXLib:
1) Create a new folder for the installation.
2) Download all the component zips to the created folder: SpTBXLib + TB2K + TB2K Patch
3) Download the MultiInstaller
4) Download the Setup.Ini file, unzip it in the folder.

The installation folder will end up with this files:

```
C:\MyInstall  
       |-  SpTBXLib.zip  
       |-  tb2k-2.2.2.zip  
       |-  TB2Kpatch-1.1.zip  
       |-  MultiInstaller.exe  
       |-  Setup.ini  
```

You are ready to install the component packages, just run the MultiInstaller, select the destination folder, and all the components will be unziped, patched, compiled and installed on the Delphi IDE.

