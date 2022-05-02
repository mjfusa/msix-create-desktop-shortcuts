# CreateManifestAndConfigForPSF.ps1
#
# Description:
# Creates modified AppxManifest.xml (AppxManifestNew.xml) and config.json to facilitate the creating of Desktop Shortcuts
# Adds AppExecutionAlias for app (Needed for Desktop Shortcuts)
# Adding additional schema definition namespaces to support AppExecutionAlias
# Replaces Executable exe with psfLauncher exe - required since Desktop Shortcuts are created via script at first time launch
# Creates config.json entries for each application in AppxManifest.xml
#
# Requires:
# AppxManifest.xml from packaged app
#
# Output:
# AppxManifestNew.xml and config.json 
# Contents of AppxManifestNew.xml and config.json file will be added to MSIX to create Desktop Shortcuts when app is run.
# CreateShortcuts.ps1, StartingScriptWrapper.ps1, and PSF binaries required. See blog post for details.

Function Export-Icons {
    <#
        .SYNOPSIS
        It's ugly but it works.
         
        .DESCRIPTION
        Export-Icons exports high-quality icons stored within .DLL and .EXE files. The function can export to a number of formats, including
        ico, bmp, png, jpg, gif, emf, exif, icon, tiff, and wmf. In addition, it can also export to a different size.
         
        This function quickly exports *all* icons stored within the resource file.
         
        .PARAMETER Path
        Path to the .dll or .exe
     
        .PARAMETER Directory
        Directory where the exports should be stored. If no directory is specified, all icons will be exported to the TEMP directory.
         
        .PARAMETER Size
        This specifies the pixel size of the exported icons. All icons will be squares, so if you want a 16x16 export, it would be -Size 16.
     
        Valid sizes are 8, 16, 24, 32, 48, 64, 96, and 128. The default is 32.
         
        .PARAMETER Type
        This is the type of file you would like to export to. The default is .ico
         
        Valid types are ico, bmp, png, jpg, gif, emf, exif, icon, tiff, and wmf. The default is ico.
         
        .NOTES
        Author: Chrissy LeMaire
        Requires: PowerShell 3.0
        Version: 1.0
        DateUpdated: 2015-Sept-16
     
        .LINK
        https://gallery.technet.microsoft.com/scriptcenter/Export-Icon-from-DLL-and-9d309047
     
        .EXAMPLE
        Export-Icons C:\windows\system32\imageres.dll
         
        Exports all icons stored witin C:\windows\system32\imageres.dll to $env:temp\icons. Creates directory if required and automatically opens output directory.
     
        .EXAMPLE
        Export-Icons -Path "C:\Program Files (x86)\VMware\Infrastructure\Virtual Infrastructure Client\Launcher\VpxClient.exe" -Size 64 -Type png -Directory C:\temp
         
        Exports the high-quality icon within VpxClient.exe to a transparent png in C:\temp\. Resizes the exported image to 64x64. Creates directory if required
        and automatically opens output directory.
         
     
        #>
        
    [CmdletBinding()] 
    Param(
        [Parameter(Mandatory = $true)]
        [string]$Path,
        [string]$Directory,
        [ValidateSet(8, 16, 24, 32, 48, 64, 96, 128)]
        [int]$Size = 32,
        [ValidateSet("ico", "bmp", "png", "jpg", "gif", "jpeg", "emf", "exif", "icon", "tiff", "wmf")]
        [string]$Type = "ico",
        [string]$iconFilename = "AppIcon"
    )
    
    BEGIN {
        
        # Thanks Thomas Levesque at http://bit.ly/1KmLgyN and darkfall http://git.io/vZxRK
        $code = '
        using System;
        using System.Drawing;
        using System.Runtime.InteropServices;
        using System.IO;
         
        namespace System {
            public class IconExtractor {
                public static Icon Extract(string file, int number, bool largeIcon) {
                    IntPtr large;
                    IntPtr small;
                    ExtractIconEx(file, number, out large, out small, 1);
                    try { return Icon.FromHandle(largeIcon ? large : small); }
                    catch { return null; }
                }
                [DllImport("Shell32.dll", EntryPoint = "ExtractIconExW", CharSet = CharSet.Unicode, ExactSpelling = true, CallingConvention = CallingConvention.StdCall)]
                private static extern int ExtractIconEx(string sFile, int iIndex, out IntPtr piLargeVersion, out IntPtr piSmallVersion, int amountIcons);
            }
        }
         
        public class PngIconConverter
        {
            public static bool Convert(System.IO.Stream input_stream, System.IO.Stream output_stream, int size, bool keep_aspect_ratio = false)
            {
                System.Drawing.Bitmap input_bit = (System.Drawing.Bitmap)System.Drawing.Bitmap.FromStream(input_stream);
                if (input_bit != null)
                {
                    int width, height;
                    if (keep_aspect_ratio)
                    {
                        width = size;
                        height = input_bit.Height / input_bit.Width * size;
                    }
                    else
                    {
                        width = height = size;
                    }
                    System.Drawing.Bitmap new_bit = new System.Drawing.Bitmap(input_bit, new System.Drawing.Size(width, height));
                    if (new_bit != null)
                    {
                        System.IO.MemoryStream mem_data = new System.IO.MemoryStream();
                        new_bit.Save(mem_data, System.Drawing.Imaging.ImageFormat.Png);
     
                        System.IO.BinaryWriter icon_writer = new System.IO.BinaryWriter(output_stream);
                        if (output_stream != null && icon_writer != null)
                        {
                            icon_writer.Write((byte)0);
                            icon_writer.Write((byte)0);
                            icon_writer.Write((short)1);
                            icon_writer.Write((short)1);
                            icon_writer.Write((byte)width);
                            icon_writer.Write((byte)height);
                            icon_writer.Write((byte)0);
                            icon_writer.Write((byte)0);
                            icon_writer.Write((short)0);
                            icon_writer.Write((short)32);
                            icon_writer.Write((int)mem_data.Length);
                            icon_writer.Write((int)(6 + 16));
                            icon_writer.Write(mem_data.ToArray());
                            icon_writer.Flush();
                            return true;
                        }
                    }
                    return false;
                }
                return false;
            }
     
            public static bool Convert(string input_image, string output_icon, int size, bool keep_aspect_ratio = false)
            {
                System.IO.FileStream input_stream = new System.IO.FileStream(input_image, System.IO.FileMode.Open);
                System.IO.FileStream output_stream = new System.IO.FileStream(output_icon, System.IO.FileMode.OpenOrCreate);
     
                bool result = Convert(input_stream, output_stream, size, keep_aspect_ratio);
     
                input_stream.Close();
                output_stream.Close();
     
                return result;
            }
        }
    '
        If (-not ("System.IconExtractor" -as [type])) {
            Add-Type -TypeDefinition $code -ReferencedAssemblies System.Drawing, System.IO -ErrorAction SilentlyContinue
        }
    }  
    PROCESS {
        switch ($type) {
            "jpg" { $type = "jpeg" }
            "icon" { $type = "ico" }
        }
            
        # Ensure file exists
        $path = Resolve-Path $path
        if ((Test-Path $path) -eq $false) { throw "$path does not exist." }
            
        # Ensure directory exists if one was specified. Otherwise, create icon directory in TEMP
        if ($directory.length -eq 0) { $directory = "$env:temp\icons" }
        if ((Test-Path $directory) -eq $false) { 
            try { New-Item -Type Directory $directory | Out-Null }
            catch { throw "Can't create $directory" }
        }
            
        # Extract
        $index = 0
        $tempfile = "$directory\tempicon.png"
        $basename = [io.path]::GetFileNameWithoutExtension($path)
            
        do {
            try { $icon = [System.IconExtractor]::Extract($path, $index, $true) } 
            catch { throw "Could not extract icon. Do you have the proper permissions?" }
            if ($icon -ne $null) {
                #Log-Write -Text "Icon extracted"
                    
                #$filepath = "$directory\$basename-$index.$type"
                $filepath = "$directory\$iconFilename"
                # Convert to bitmap, otherwise it's ugly
                $bmp = $icon.ToBitmap()
                    
                try { 
                    if ($type -eq "ico") {
                        $bmp.Save($tempfile, "png")
                        [PngIconConverter]::Convert($tempfile, $filepath, $size, $true) | Out-Null
                        # Keep remove-item from complaining about weird directories
                        cmd /c del $tempfile
                        Log-Write -Text "Icon converted to png for resizing."
                    }
                    else {                    
                        if ($bmp.Width -ne $size) {
                            # Needs to be resized
                            $newbmp = New-Object System.Drawing.Bitmap($size, $size)
                            $graph = [System.Drawing.Graphics]::FromImage($newbmp)
                            
                            # Make it transparent
                            $graph.clear([System.Drawing.Color]::Transparent)
                            $graph.DrawImage($bmp, 0, 0, $size, $size)
                             
                            #save to file
                            $newbmp.Save($filepath, $type)
                            $newbmp.Dispose()
                        }
                        else { $bmp.Save($filepath, $type) }
                        $bmp.Dispose()
                    }
                    $icon.Dispose()                
                    $index++
                }
                catch { throw "Could not convert icon." }
    
            }
        } while ($icon -ne $null)
            
        if ($index -eq 0) { Write-Error "No icons to extract :(" }
    }
}




# Create config.json 
$scriptJson = @{
    scriptPath = 'CreateShortcuts.ps1'
}

$applications = @{
    id = ''
    executable = ''
    arguments = ''
    startScript = $scriptJson
}
$configjsonArray = [System.Collections.ArrayList]@()

# Get AppIds and Executable names from manifest to write to config.json
[xml]$manifest = get-content "AppxManifest.xml"
$appsInManifest= $manifest.Package.Applications.Application
foreach ($app in $appsInManifest) {
    if ($app.VisualElements.AppListEntry -eq "none") {
        continue
    }
    $object = new-object psobject -Property $applications
    $object.id= $app.id;
    $object.executable = $app.Executable;
    $res=$configjsonArray.Add($object);
}

# Create config.json file
# config.json Application arguments
$header="{
    `"applications`":["

$footer="]}"

$header | Out-File  ($pwd.Path + "\config.json");
$configjsonArray | ConvertTo-Json | Out-File -Append ($pwd.Path + "\config.json");
# Write config.json
$footer | Out-File -Append ($pwd.Path + "\config.json");

# Add namespaces needed by AppExecutionAlias
$nsm = New-Object System.Xml.XmlNamespaceManager($manifest.nametable)
$nsList = @( ("uap3", "http://schemas.microsoft.com/appx/manifest/uap/windows10/3"), ("desktop", "http://schemas.microsoft.com/appx/manifest/desktop/windows10") )
$ignoreNS=$manifest.Package.IgnorableNamespaces;

foreach ($ns in $nsList) {
    $nsm.AddNamespace($ns[0], $ns[1])
    $manifest.Package.SetAttribute("xmlns:"+$ns[0], $ns[1])
    $ignoreNS = $ignoreNS + " " + $ns[0]
}
$manifest.Package.RemoveAttribute("IgnorableNamespaces")
$manifest.Package.SetAttribute("IgnorableNamespaces", $ignoreNS)

# Determine which psf launcher is needed
$arch = $manifest.Package.Identity.ProcessorArchitecture 
if ($arch -eq "x64") {
    $psfLauncher='psflauncher64.exe'
} else {
    $psfLauncher='psflauncher32.exe'    
}

# Add appExecutionAlias for each exe in the manifest, except if AppListEntry='None'
# Create appExecutionAlias node
foreach ($app in $manifest.Package.Applications.Application) {
    $el=$manifest.CreateElement("uap3:Extension", $nsm.LookupNamespace("uap3"))
    $el.SetAttribute("Category", "windows.appExecutionAlias")
    $el.SetAttribute("Executable", $psfLauncher)
    $el.SetAttribute("EntryPoint", "Windows.FullTrustApplication")
    
    $aliasEl=$manifest.CreateElement("uap3:AppExecutionAlias", $nsm.LookupNamespace("uap3"))
    $appaliasEl=$manifest.CreateElement("desktop:ExecutionAlias", $nsm.LookupNamespace("desktop"))
    # Note: Alias has same name as EXE
    $res = $appaliasEl.SetAttribute("Alias",  $($(Split-Path $app.Executable -Leaf).ToLower())) 
    
    $res= $aliasEl.AppendChild($appaliasEl)
    $res= $el.AppendChild($aliasEl) 

    if ($null -eq $app.Extensions) {
        $extensions= $manifest.CreateElement("Extensions")
        $ext= $manifest.Package.Applications.Application.AppendChild($extensions)
        $res = $ext.AppendChild($el)
        $extensions.SetAttribute("xmlns", "http://schemas.microsoft.com/appx/manifest/foundation/windows10")
    } else {
        $app.Extensions.AppendChild($el)
    }
}

# Set Executables in manifest to PSFLauncher
foreach ($app in $manifest.Package.Applications.Application) {
    $app.SetAttribute("Executable", $psfLauncher);
}

# Write new manifest file
$manifest.Save($pwd.Path + "\AppxManifestNew.xml");

#Create Icons
$desinationIconPath = $($pwd.Path + "\Icons")
$res = Test-Path $desinationIconPath
if ($res -ne $true)
{
    New-Item -Path $desinationIconPath -ItemType Directory
}

Function ExtractIcon {
    Param ( 
        [Parameter(Mandatory = $true)]
        [string]$exePath,
        [string]$destinationFolder,
        [string]$iconName
    )
    $destIcon = $iconName + ".ico"
    Export-Icons $exePath $destinationFolder  32 "ico" $destIcon
}

$index = 0;
Foreach ($app in $configjsonArray) {
    $appPath = ($pwd.Path + "\" + $app.Executable)
    try {
        ExtractIcon -exePath ( $appPath ) -destinationFolder ( $desinationIconPath ) -iconName ( $appsInManifest.VisualElements.DisplayName )
    } catch {
        Write-Host $Error[0]
    }
    if (Test-Path($($desinationIconPath + "\" + $appsInManifest.VisualElements.DisplayName + ".ico")) ) {
        Write-Host $("Icon successfully created: " + $desinationIconPath + "\" + $appsInManifest.VisualElements.DisplayName + ".ico")
    } else {
        Write-Host $("Icon NOT created: " + $desinationIconPath + "\" + $appsInManifest.VisualElements.DisplayName + ".ico")
    }
}
