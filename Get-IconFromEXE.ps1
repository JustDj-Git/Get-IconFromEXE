<#
    .SYNOPSIS
    	Get-IconFromEXE extracts the icon image
	
    .DESCRIPTION
    	Get-IconFromEXE extracts the icon image from an exe file and saves it in the same directory as the .exe file
	
    .PARAMETER Path
    	Full path to the .exe file
	
    .EXAMPLE
    	Get-IconFromEXE -Path 'C:\location\test.exe'
    	Extract the image from test.exe and saves it to the icon.ico file in folder 'C:\location\'
	
    .EXAMPLE
    	Get-IconFromEXE -Path 'C:\location\test.exe' -ImageFormat bmp
    	Extract the image from test.exe and saves it to the icon.bmp file in folder 'C:\location\'
	
    .EXAMPLE
    	'C:\location\test.exe', 'C:\location2\test2.exe' | Get-IconFromEXE -ImageFormat png
    	Extract the image from test.exe and test2.exe and saves it to the icon.bmp file in folder 'C:\location\' and 'C:\location2\'
	
    .EXAMPLE
    	Get-IconFromEXE -Path 'C:\location'
    	Extract the image from any first .exe file and saves it to the icon.ico file in folder 'C:\location\'

    .NOTES
    	Written by JustDj (justdj.ca@gmail.com)
    	-If 'Path' is not set - opens system FileDialog
    	-If 'ImageFormat' is not set - will be used .ico by default
    	-If 'Path' does not contains exe file - will be selected first one in folder
	
    .LINK
    	https://github.com/JustDj-Git/Get-IconFromEXE
#>
function Get-IconFromEXE {
	param
	(
		[Parameter(ValueFromPipeline)]
		[string[]]$Path,
		[ValidateSet('ico', 'bmp', 'png', 'jpg')]
		$ImageFormat = 'ico',
		[Int]$Index = 0,
		[Switch]$LargeIcon
	)
	
	begin {
		#for ico converver
		$TypeDefinition = @'
using System.Drawing;
using System.Drawing.Imaging;
using System.IO;
using System.Collections.Generic;
using System.Drawing.Drawing2D;

/// <summary>
/// Adapted from this gist: https://gist.github.com/darkfall/1656050
/// Provides helper methods for imaging
/// </summary>
public static class ImagingHelper
{
    /// <summary>
    /// Converts a PNG image to a icon (ico) with all the sizes windows likes
    /// </summary>
    /// <param name="inputBitmap">The input bitmap</param>
    /// <param name="output">The output stream</param>
    /// <returns>Wether or not the icon was succesfully generated</returns>
    public static bool ConvertToIcon(Bitmap inputBitmap, Stream output)
    {
        if (inputBitmap == null)
            return false;

        int[] sizes = new int[] { 256, 48, 32, 16 };

        // Generate bitmaps for all the sizes and toss them in streams
        List<MemoryStream> imageStreams = new List<MemoryStream>();
        foreach (int size in sizes)
        {
            Bitmap newBitmap = ResizeImage(inputBitmap, size, size);
            if (newBitmap == null)
                return false;
            MemoryStream memoryStream = new MemoryStream();
            newBitmap.Save(memoryStream, ImageFormat.Png);
            imageStreams.Add(memoryStream);
        }

        BinaryWriter iconWriter = new BinaryWriter(output);
        if (output == null || iconWriter == null)
            return false;

        int offset = 0;

        // 0-1 reserved, 0
        iconWriter.Write((byte)0);
        iconWriter.Write((byte)0);

        // 2-3 image type, 1 = icon, 2 = cursor
        iconWriter.Write((short)1);

        // 4-5 number of images
        iconWriter.Write((short)sizes.Length);

        offset += 6 + (16 * sizes.Length);

        for (int i = 0; i < sizes.Length; i++)
        {
            // image entry 1
            // 0 image width
            iconWriter.Write((byte)sizes[i]);
            // 1 image height
            iconWriter.Write((byte)sizes[i]);

            // 2 number of colors
            iconWriter.Write((byte)0);

            // 3 reserved
            iconWriter.Write((byte)0);

            // 4-5 color planes
            iconWriter.Write((short)0);

            // 6-7 bits per pixel
            iconWriter.Write((short)32);

            // 8-11 size of image data
            iconWriter.Write((int)imageStreams[i].Length);

            // 12-15 offset of image data
            iconWriter.Write((int)offset);

            offset += (int)imageStreams[i].Length;
        }

        for (int i = 0; i < sizes.Length; i++)
        {
            // write image data
            // png data must contain the whole png data file
            iconWriter.Write(imageStreams[i].ToArray());
            imageStreams[i].Close();
        }

        iconWriter.Flush();

        return true;
    }

    /// <summary>
    /// Converts a PNG image to a icon (ico)
    /// </summary>
    /// <param name="input">The input stream</param>
    /// <param name="output">The output stream</param
    /// <returns>Wether or not the icon was succesfully generated</returns>
    public static bool ConvertToIcon(Stream input, Stream output)
    {
        Bitmap inputBitmap = (Bitmap)Bitmap.FromStream(input);
        return ConvertToIcon(inputBitmap, output);
    }

    /// <summary>
    /// Converts a PNG image to a icon (ico)
    /// </summary>
    /// <param name="inputPath">The input path</param>
    /// <param name="outputPath">The output path</param>
    /// <returns>Wether or not the icon was succesfully generated</returns>
    public static bool ConvertToIcon(string inputPath, string outputPath)
    {
        using (FileStream inputStream = new FileStream(inputPath, FileMode.Open))
        using (FileStream outputStream = new FileStream(outputPath, FileMode.OpenOrCreate))
        {
            return ConvertToIcon(inputStream, outputStream);
        }
    }



    /// <summary>
    /// Converts an image to a icon (ico)
    /// </summary>
    /// <param name="inputImage">The input image</param>
    /// <param name="outputPath">The output path</param>
    /// <returns>Wether or not the icon was succesfully generated</returns>
    public static bool ConvertToIcon(Image inputImage, string outputPath)
    {
        using (FileStream outputStream = new FileStream(outputPath, FileMode.OpenOrCreate))
        {
            return ConvertToIcon(new Bitmap(inputImage), outputStream);
        }
    }


    /// <summary>
    /// Resize the image to the specified width and height.
    /// Found on stackoverflow: https://stackoverflow.com/questions/1922040/resize-an-image-c-sharp
    /// </summary>
    /// <param name="image">The image to resize.</param>
    /// <param name="width">The width to resize to.</param>
    /// <param name="height">The height to resize to.</param>
    /// <returns>The resized image.</returns>
    public static Bitmap ResizeImage(Image image, int width, int height)
    {
        var destRect = new Rectangle(0, 0, width, height);
        var destImage = new Bitmap(width, height);

        destImage.SetResolution(image.HorizontalResolution, image.VerticalResolution);

        using (var graphics = Graphics.FromImage(destImage))
        {
            graphics.CompositingMode = CompositingMode.SourceCopy;
            graphics.CompositingQuality = CompositingQuality.HighQuality;
            graphics.InterpolationMode = InterpolationMode.HighQualityBicubic;
            graphics.SmoothingMode = SmoothingMode.HighQuality;
            graphics.PixelOffsetMode = PixelOffsetMode.HighQuality;

            using (var wrapMode = new ImageAttributes())
            {
                wrapMode.SetWrapMode(WrapMode.TileFlipXY);
                graphics.DrawImage(image, destRect, 0, 0, image.Width, image.Height, GraphicsUnit.Pixel, wrapMode);
            }
        }

        return destImage;
    }
}
'@
		
		Add-Type -TypeDefinition $TypeDefinition -ReferencedAssemblies 'System.Drawing'
		
		Add-Type -Namespace Win32API -Name Icon -MemberDefinition @'
[DllImport("Shell32.dll", SetLastError=true)]
public static extern int ExtractIconEx(string lpszFile, int nIconIndex, out IntPtr phiconLarge, out IntPtr phiconSmall, int nIcons);

[DllImport("gdi32.dll", SetLastError=true)]
public static extern bool DeleteObject(IntPtr hObject);
'@
		
		#avoid Path var
		[string[]]$exePath = $Path
		$null = Add-Type -AssemblyName System.Drawing
		$ErrorActionPreference = 'Stop'
	}
	
	process {
		try {
			if (!($exePath)) {
				Add-Type -AssemblyName System.Windows.Forms
				$OpenFileDialog = New-Object -TypeName System.Windows.Forms.OpenFileDialog
				$OpenFileDialog.restoreDirectory = $True
				$OpenFileDialog.title = 'Select an EXE File'
				$OpenFileDialog.filter = 'Executable files (*.exe)|*.exe'
				if (($OpenFileDialog.ShowDialog() -eq 'OK')) {
					$exePath = $OpenFileDialog.FileName
				} else {
					Write-Host 'Do it again' -ForegroundColor Red
					return
				}
			}
			
			function GetICO{
				$large, $small = 0, 0
				$null = [Win32API.Icon]::ExtractIconEx($($resolved_path.FullName), $Index, [ref]$large, [ref]$small, 1)
				$handle = if ($LargeIcon) { $large } else { $small }
				if ($handle) { $icon = [System.Drawing.Icon]::FromHandle($handle) }
				$large, $small, $handle | Where-Object { $_ } | ForEach-Object { [Win32API.Icon]::DeleteObject($_) } | Out-Null
				return $icon
			}
			
			foreach ($i in $exePath) {
				if (Test-Path $i){
					$resolved_path = Get-ChildItem -Path $i -Filter '*.exe' -ErrorAction SilentlyContinue | Where-Object { $_.Name -like "*$(Split-Path -Path $i -Leaf)*" }
					if (!($resolved_path)) {
						$resolved_path = Get-ChildItem -Path $i -Filter '*.exe' -ErrorAction SilentlyContinue | Select-Object -First 1
					}
					
					if ($resolved_path) {
						Write-Host "Extracting icon from $($resolved_path.Name)" -ForegroundColor Yellow
						$icon = GetICO
						if ($ImageFormat -eq 'ico') {
							$out = "$($resolved_path.Directory.FullName)" + '\icon.png'
							$icon.ToBitmap().Save($out)
							$null = [ImagingHelper]::ConvertToIcon($out, "$($resolved_path.Directory.FullName)" + '\icon.ico')
							if(Test-Path $out){Remove-Item $out -Force -ErrorAction SilentlyContinue}
						} else {
							$out = "$($resolved_path.Directory.FullName)" + '\icon.' + "$ImageFormat"
							$icon.ToBitmap().Save($out)
						}
					} else { Write-host "No exe files in path $i" -ForegroundColor Red }
				} else { Write-host "Path is not exist $i" -ForegroundColor Red }
			} #foreach end
		} catch {
			Write-Host "`n$_" -ForegroundColor Red
			Write-Host "`n$($_.ScriptStackTrace)`n" -ForegroundColor Red
		}
	}
	end { Write-Host 'Done' -ForegroundColor Green }
}

Get-IconFromEXE -Path 'D:\Progs\GoldWave', 'D:\Progs\test\TankIconMaker' -ImageFormat png