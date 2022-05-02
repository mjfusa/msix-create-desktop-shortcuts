# Sample - MSIX Create Desktop Shortcuts

See the Microsoft TechCommunity post here XXXLINKXXX for more information:

## In this repo

```ScriptInPackage\CreateShortcuts.ps1``` - First time run script that creates desktop icons for application.

```CreateManifestAndConfigForPSF.ps1``` - 'Prep script'. Creates new AppxManifest files based on existing manifest, config.json for Package Support Framework, and Icons extracted from desktop application.

```Sample-Manifest-Config``` - Sample ```AppxManifestNew.xml``` and ```config.json``` files.

```start``` - HelloWorld.msix. An application you can test the following instructions against, before doing this with your own app.

```complete``` - Hello.MSIX.World_1.0.1.0_x64__5ncaqh95vjk2e.msix. For your reference. HelloWorld.MSIX with the changes necessary to create shortcuts.

## Prerequisites

You'll need the following:

1) ```CreateShortcuts.ps1```. Clone this repository: https://github.com/mjfusa/msix-create-desktop-shortcuts. The script is in the ```ScriptInPackage``` folder.
2) An icon embedded in your application's EXE. This will be used to create the shortcut icon.
3) Package Support Framework binaries. Here are the steps to download:

```PowerShell
# PowerShell - Download PsdBinaries 2.0 into the PsfBinaries folder
md PsfBinaires
cd PsfBinaires
wget https://github.com/microsoft/MSIX-PackageSupportFramework/releases/download/v2.0/PSFBinaries.zip -OutFile PSFBinaries.zip
Expand-Archive .\PSFBinaries.zip
```

4) **MSIX Packaging Tool**. Install from the Microsoft Store [here](https://www.microsoft.com/en-us/p/msix-packaging-tool/9n5lw3jbcxkf).
5) An MSIX Packaged app.
## Preparing your MSIX
### Create Icons, AppxmanifestNew.xml and config.json.

You will need to start with an existing MSIX packaged app.

1) Extract the contents of your MSIX package to a temporary directory. You can use the ```MSIX Packaging Tool``` to do this.

2) Change to the temporary directory where you unpackaged the application.

3) Run the script ```CreateManifestAndConfigForPSF.ps1``` in the same folder where you have saved the extracted the application.

4) After running the script, the icon(s) are extracted, converted to **ico** files and saved to the 'Icons' folder. Also ```AppxmanifestNew.xml``` and ```config.json``` are created. ```config.json``` is required by the Package Support Framework. You will replace the existing manifest with the contents of ```AppxmanifestNew.xml```.

### Update your MSIX Package with the MSIX Packaging Tool

#### Update the manifest

1. Open your MSIX package for editing in the **MSIX Packaging tool** (MPT)

2. In the 'Manifest file' section click on 'Open File'. The manifest will be opened into Notepad for editing.

3. Open ```AppmanifestNew.xml``` created in the section above in another editor and copy the entire contents to the clipboard.

4. Go back Notepad. Delete all of the contents and paste in the contents you copied in the previous step.

5. Notepad should now have the contents of ```AppmanifestNew.xml```. Save the file and close notepad.

6. Verify in **MPT** that you have no errors in the 'Manifest file' section.

#### Add config.json

1. Click on 'Package files' in **MPT**

2. Right click on ```Package```

3. Click ```Add File``` and add the ```config.json``` file that was created in the [Create Icons, AppxmanifestNew.xml and config](#create-appxmanifestnewxml-and-configjson.json) section above.

#### <a name="add-the-icons"></a>Add the Icons

1. Click on 'Package files' in **MPT**

2. Right click on ```Package```

3. Click on ```New folder```. Type ```Icons```. Click **Save**

4. Right click on ```Icons```

5. Click ```Add File``` and add the png files that were created in the section [Create Icons, AppxmanifestNew.xml and config](#create-appxmanifestnewxml-and-configjson.json) above.

#### Add the Package Support Framework (PSF) Binaries and CreateShortcuts.ps1 script

Note: For this step you will need the PSF Binaries and the ```CreateShortcuts.ps1``` script included in the GitHub repository. See **Prerequisites**. 

1. Go back to MPT

2. Click on **'Package files'** in MPT

3. Right click on ```Package```

4. Click 'Add File' and add the following files from the ```PSFBinaries\bin``` folder:

```txt
PsfLauncher32.exe
PsfLauncher64.exe
PsfRunDll32.exe
PsfRunDll64.exe
PsfRuntime32.dll
PsfRuntime64.dll
StartingScriptWrapper.ps1
```

The ```StartingScriptWrapper.ps1``` is required to run PowerShell scripts from PSF.

#### Add the script ```CreateShortcuts.ps1```

1. Go back to MPT
2. Click on 'Package files' in MPT
3. Right click on ```Package```
4. Click 'Add File' and add the file ```CreateShortcuts.ps1```

#### Save the MSIX package

You have added all of the necessary files and can now save and sign the MSIX.

1. Go back to **MPT**
2. Provide the code signing certificate.
3. Click **Save** and click **Yes** to increment the version number.
4. The MSIX is updated with create shortcut support, signed and ready for testing.
