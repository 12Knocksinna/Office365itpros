# Convert-TeamsBackgroundFilesV2.PS1
# Script to convert custom background images used with Teams V1 so that they can be used with the
# V2 client.
# https://github.com/12Knocksinna/Office365itpros/blob/master/Convert-TeamsBackgroundFilesV2.PS1

<#
.SYNOPSIS
   Resize an image 
.DESCRIPTION
   Resize an image based on a new given height or width or a single dimension and a maintain ratio flag. 
   The execution of this CmdLet creates a new file named "OriginalName_resized" and maintains the original
   file extension
.PARAMETER Width
   The new width of the image. Can be given alone with the MaintainRatio flag
.PARAMETER Height
   The new height of the image. Can be given alone with the MaintainRatio flag
.PARAMETER ImagePath
   The path to the image being resized
.PARAMETER MaintainRatio
   Maintain the ratio of the image by setting either width or height. Setting both width and height and also this parameter
   results in an error
.PARAMETER Percentage
   Resize the image *to* the size given in this parameter. It's imperative to know that this does not resize by the percentage but to the percentage of
   the image.
.PARAMETER SmoothingMode
   Sets the smoothing mode. Default is HighQuality.
.PARAMETER InterpolationMode
   Sets the interpolation mode. Default is HighQualityBicubic.
.PARAMETER PixelOffsetMode
   Sets the pixel offset mode. Default is HighQuality.
.EXAMPLE
   Resize-Image -Height 45 -Width 45 -ImagePath "Path/to/image.jpg"
.EXAMPLE
   Resize-Image -Height 45 -MaintainRatio -ImagePath "Path/to/image.jpg"
.EXAMPLE
   #Resize to 50% of the given image
   Resize-Image -Percentage 50 -ImagePath "Path/to/image.jpg"
.NOTES
   Written By: 
   Christopher Walker https://gist.github.com/someshinyobject/617bf00556bc43af87cd
#>
Function Resize-Image() {
    [CmdLetBinding(
        SupportsShouldProcess=$true, 
        PositionalBinding=$false,
        ConfirmImpact="Medium",
        DefaultParameterSetName="Absolute"
    )]
    Param (
        [Parameter(Mandatory=$True)]
        [ValidateScript({
            $_ | ForEach-Object {
                Test-Path $_
            }
        })][String[]]$ImagePath,
        [Parameter(Mandatory=$False)][Switch]$MaintainRatio,
        [Parameter(Mandatory=$False, ParameterSetName="Absolute")][Int]$Height,
        [Parameter(Mandatory=$False, ParameterSetName="Absolute")][Int]$Width,
        [Parameter(Mandatory=$False, ParameterSetName="Percent")][Double]$Percentage,
        [Parameter(Mandatory=$False)][System.Drawing.Drawing2D.SmoothingMode]$SmoothingMode = "HighQuality",
        [Parameter(Mandatory=$False)][System.Drawing.Drawing2D.InterpolationMode]$InterpolationMode = "HighQualityBicubic",
        [Parameter(Mandatory=$False)][System.Drawing.Drawing2D.PixelOffsetMode]$PixelOffsetMode = "HighQuality",
        [Parameter(Mandatory=$False)][String]$NameModifier = "resized"
    )
    Begin {
        If ($Width -and $Height -and $MaintainRatio) {
            Throw "Absolute Width and Height cannot be given with the MaintainRatio parameter."
        }
 
        If (($Width -xor $Height) -and (-not $MaintainRatio)) {
            Throw "MaintainRatio must be set with incomplete size parameters (Missing height or width without MaintainRatio)"
        }
 
        If ($Percentage -and $MaintainRatio) {
            Write-Warning "The MaintainRatio flag while using the Percentage parameter does nothing"
        }
    }
    Process {
        Add-Type -AssemblyName 'System.Drawing'
        ForEach ($Image in $ImagePath) {
            $Path = (Resolve-Path $Image).Path
            $Dot = $Path.LastIndexOf(".")

            #Add name modifier (OriginalName_{$NameModifier}.jpg)
            $OutputPath = $Path.Substring(0,$Dot) + "_" + $NameModifier + $Path.Substring($Dot,$Path.Length - $Dot)
            
            $OldImage = New-Object -TypeName System.Drawing.Bitmap -ArgumentList $Path
            # Grab these for use in calculations below. 
            $OldHeight = $OldImage.Height
            $OldWidth = $OldImage.Width
 
            If ($MaintainRatio) {
                $OldHeight = $OldImage.Height
                $OldWidth = $OldImage.Width
                If ($Height) {
                    $Width = $OldWidth / $OldHeight * $Height
                }
                If ($Width) {
                    $Height = $OldHeight / $OldWidth * $Width
                }
            }
 
            If ($Percentage) {
                $Product = ($Percentage / 100)
                $Height = $OldHeight * $Product
                $Width = $OldWidth * $Product
            }

            $Bitmap = New-Object -TypeName System.Drawing.Bitmap -ArgumentList $Width, $Height
            $NewImage = [System.Drawing.Graphics]::FromImage($Bitmap)
             
            #Retrieving the best quality possible
            $NewImage.SmoothingMode = $SmoothingMode
            $NewImage.InterpolationMode = $InterpolationMode
            $NewImage.PixelOffsetMode = $PixelOffsetMode
            $NewImage.DrawImage($OldImage, $(New-Object -TypeName System.Drawing.Rectangle -ArgumentList 0, 0, $Width, $Height))

            If ($PSCmdlet.ShouldProcess("Resized image based on $Path", "save to $OutputPath")) {
                $Bitmap.Save($OutputPath)
            }
            
            $Bitmap.Dispose()
            $NewImage.Dispose()
            $OldImage.Dispose()
        }
    }
}

# Start of processing to convert Teams background images

# First, identify the target directory where Teams V2 stores its background images. Each
# image has two files - a full version and a thumbnail.
$TargetDirectory = ($Env:appdata -Split("\\Roaming"))[0] + "\Local\Packages\MSTeams_8wekyb3d8bbwe\LocalCache\Microsoft\MSTeams\Backgrounds\Uploads"
# Now find the folder where Teams V1 stored its background files
$TeamsV1Directory = $Env:appdata + "\Microsoft\Teams\Backgrounds\Uploads"

# Check that the two folders are available
If (!(Test-Path -Path $TargetDirectory)) {
    New-Item -Path $TargetDirectory -ItemType Directory
}

If (!(Test-Path $TeamsV1Directory)) {
    Write-Host ("The Teams V1 folder for background images is not on this PC - exiting!" -f $TeamsV1Directory)
    break
}

# Get the set of files that we can convert
Set-Location -Path $TeamsV1Directory
[array]$Files = Get-ChildItem "*.jpg" | Select-Object -ExpandProperty Name
[int]$i =0
# If we can find some JPG files to convert, go ahead and create new files
If ($Files) {
 # Remove thumb nails and only process full definition files   
 $Files = $Files | Where-Object {$_ -notlike "*_thumb.jpg*"}
 ForEach ($File in $Files) {
    Write-Host ("Processing {0}..." -f $File)
    $Name = $File.Split(".")[0]
    $OutputFile = $Name + "_resized.jpg"
    # Create 1920x1080 version of file
    Resize-image -ImagePath $File -height 1080 -Width 1920
    If ($OutputFile) {
        $Guid = New-Guid
        $FullFileName = ("{0}.jpg" -f $Guid)
        Rename-Item $OutputFile $FullFileName
    }
    # Now create the thumbnail image
    $OutputFile = $Name + "_resized.jpg"
    Resize-image -ImagePath $File -height 158 -Width 220
    If ($OutputFile) {
        $ThumbFileName = ("{0}_thumb.jpg" -f $Guid)
        Rename-Item $OutputFile $ThumbFileName
    }
    $CopyFiles = ("{0}*.jpg" -f $Guid)
    # Copy files to the Teams V2 folder and clean up by removing them from the V1 folder
    Copy-Item $CopyFiles $TargetDirectory
    Remove-Item $CopyFiles
    $i++
 } # End ForEach Files
} # End if Files
Write-Host ("All done. {0} files processed and copied to the Teams V2 folder. You should be able to use these files as background images in Teams V2 meetings" -f $i)
Set-Location -Path $TargetDirectory
Get-ChildItem "*.jpg"
Set-Location -Path "C:\"

# An example script used to illustrate a concept. More information about the topic can be found in the Office 365 for IT Pros eBook https://gum.co/O365IT/
# and/or a relevant article on https://office365itpros.com or https://www.practical365.com. See our post about the Office 365 for IT Pros repository 
# https://office365itpros.com/office-365-github-repository/ for information about the scripts we write.

# Do not use our scripts in production until you are satisfied that the code meets the needs of your organization. Never run any code downloaded from 
# the Internet without first validating the code in a non-production environment.
