#
# Powershell Version: 3.0
# Language:       English
# Platform:       Win x86/x64
# Author:         gopekanna@gmail.com
#
# Script Function:
#   Voice call, Video call, Desktop sharing and Mut/Unmute the microphone witha a particular contact and capture the test evidence as screenshot including the HTML test report
#

# Inbuilt function to get the screenshot of the desktop/laptop screen
Function Take-ScreenShot { 
    <#   
.SYNOPSIS   
    Used to take a screenshot of the desktop or the active window.  
.DESCRIPTION   
    Used to take a screenshot of the desktop or the active window and save to an image file if needed. 
.PARAMETER screen 
    Screenshot of the entire screen 
.PARAMETER activewindow 
    Screenshot of the active window 
.PARAMETER file 
    Name of the file to save as. Default is image.bmp 
.PARAMETER imagetype 
    Type of image being saved. Can use JPEG,BMP,PNG. Default is bitmap(bmp)   
.PARAMETER print 
    Sends the screenshot directly to your default printer       
.INPUTS 
.OUTPUTS     
.NOTES   
    Name: Take-ScreenShot 
    Author: Boe Prox 
    DateCreated: 07/25/2010      
.EXAMPLE   
    Take-ScreenShot -activewindow 
    Takes a screen shot of the active window         
.EXAMPLE   
    Take-ScreenShot -Screen 
    Takes a screenshot of the entire desktop 
.EXAMPLE   
    Take-ScreenShot -activewindow -file "C:\image.bmp" -imagetype bmp 
    Takes a screenshot of the active window and saves the file named image.bmp with the image being bitmap 
.EXAMPLE   
    Take-ScreenShot -screen -file "C:\image.png" -imagetype png     
    Takes a screenshot of the entire desktop and saves the file named image.png with the image being png 
.EXAMPLE   
    Take-ScreenShot -Screen -print 
    Takes a screenshot of the entire desktop and sends to a printer 
.EXAMPLE   
    Take-ScreenShot -ActiveWindow -print 
    Takes a screenshot of the active window and sends to a printer     
#>   
#Requires -Version 2 
        [cmdletbinding( 
                SupportsShouldProcess = $True, 
                DefaultParameterSetName = "screen", 
                ConfirmImpact = "low" 
        )] 
Param ( 
       [Parameter( 
            Mandatory = $False, 
            ParameterSetName = "screen", 
            ValueFromPipeline = $True)] 
            [switch]$screen, 
       [Parameter( 
            Mandatory = $False, 
            ParameterSetName = "window", 
            ValueFromPipeline = $False)] 
            [switch]$activewindow, 
       [Parameter( 
            Mandatory = $False, 
            ParameterSetName = "", 
            ValueFromPipeline = $False)] 
            [string]$file,  
       [Parameter( 
            Mandatory = $False, 
            ParameterSetName = "", 
            ValueFromPipeline = $False)] 
            [string] 
            [ValidateSet("bmp","jpeg","png")] 
            $imagetype = "bmp", 
       [Parameter( 
            Mandatory = $False, 
            ParameterSetName = "", 
            ValueFromPipeline = $False)] 
            [switch]$print                        
        
) 
# C# code 
$code = @' 
using System; 
using System.Runtime.InteropServices; 
using System.Drawing; 
using System.Drawing.Imaging; 
namespace ScreenShotDemo 
{ 
  /// <summary> 
  /// Provides functions to capture the entire screen, or a particular window, and save it to a file. 
  /// </summary> 
  public class ScreenCapture 
  { 
    /// <summary> 
    /// Creates an Image object containing a screen shot the active window 
    /// </summary> 
    /// <returns></returns> 
    public Image CaptureActiveWindow() 
    { 
      return CaptureWindow( User32.GetForegroundWindow() ); 
    } 
    /// <summary> 
    /// Creates an Image object containing a screen shot of the entire desktop 
    /// </summary> 
    /// <returns></returns> 
    public Image CaptureScreen() 
    { 
      return CaptureWindow( User32.GetDesktopWindow() ); 
    }     
    /// <summary> 
    /// Creates an Image object containing a screen shot of a specific window 
    /// </summary> 
    /// <param name="handle">The handle to the window. (In windows forms, this is obtained by the Handle property)</param> 
    /// <returns></returns> 
    private Image CaptureWindow(IntPtr handle) 
    { 
      // get te hDC of the target window 
      IntPtr hdcSrc = User32.GetWindowDC(handle); 
      // get the size 
      User32.RECT windowRect = new User32.RECT(); 
      User32.GetWindowRect(handle,ref windowRect); 
      int width = windowRect.right - windowRect.left; 
      int height = windowRect.bottom - windowRect.top; 
      // create a device context we can copy to 
      IntPtr hdcDest = GDI32.CreateCompatibleDC(hdcSrc); 
      // create a bitmap we can copy it to, 
      // using GetDeviceCaps to get the width/height 
      IntPtr hBitmap = GDI32.CreateCompatibleBitmap(hdcSrc,width,height); 
      // select the bitmap object 
      IntPtr hOld = GDI32.SelectObject(hdcDest,hBitmap); 
      // bitblt over 
      GDI32.BitBlt(hdcDest,0,0,width,height,hdcSrc,0,0,GDI32.SRCCOPY); 
      // restore selection 
      GDI32.SelectObject(hdcDest,hOld); 
      // clean up 
      GDI32.DeleteDC(hdcDest); 
      User32.ReleaseDC(handle,hdcSrc); 
      // get a .NET image object for it 
      Image img = Image.FromHbitmap(hBitmap); 
      // free up the Bitmap object 
      GDI32.DeleteObject(hBitmap); 
      return img; 
    } 
    /// <summary> 
    /// Captures a screen shot of the active window, and saves it to a file 
    /// </summary> 
    /// <param name="filename"></param> 
    /// <param name="format"></param> 
    public void CaptureActiveWindowToFile(string filename, ImageFormat format) 
    { 
      Image img = CaptureActiveWindow(); 
      img.Save(filename,format); 
    } 
    /// <summary> 
    /// Captures a screen shot of the entire desktop, and saves it to a file 
    /// </summary> 
    /// <param name="filename"></param> 
    /// <param name="format"></param> 
    public void CaptureScreenToFile(string filename, ImageFormat format) 
    { 
      Image img = CaptureScreen(); 
      img.Save(filename,format); 
    }     
    
    /// <summary> 
    /// Helper class containing Gdi32 API functions 
    /// </summary> 
    private class GDI32 
    { 
       
      public const int SRCCOPY = 0x00CC0020; // BitBlt dwRop parameter 
      [DllImport("gdi32.dll")] 
      public static extern bool BitBlt(IntPtr hObject,int nXDest,int nYDest, 
        int nWidth,int nHeight,IntPtr hObjectSource, 
        int nXSrc,int nYSrc,int dwRop); 
      [DllImport("gdi32.dll")] 
      public static extern IntPtr CreateCompatibleBitmap(IntPtr hDC,int nWidth, 
        int nHeight); 
      [DllImport("gdi32.dll")] 
      public static extern IntPtr CreateCompatibleDC(IntPtr hDC); 
      [DllImport("gdi32.dll")] 
      public static extern bool DeleteDC(IntPtr hDC); 
      [DllImport("gdi32.dll")] 
      public static extern bool DeleteObject(IntPtr hObject); 
      [DllImport("gdi32.dll")] 
      public static extern IntPtr SelectObject(IntPtr hDC,IntPtr hObject); 
    } 
 
    /// <summary> 
    /// Helper class containing User32 API functions 
    /// </summary> 
    private class User32 
    { 
      [StructLayout(LayoutKind.Sequential)] 
      public struct RECT 
      { 
        public int left; 
        public int top; 
        public int right; 
        public int bottom; 
      } 
      [DllImport("user32.dll")] 
      public static extern IntPtr GetDesktopWindow(); 
      [DllImport("user32.dll")] 
      public static extern IntPtr GetWindowDC(IntPtr hWnd); 
      [DllImport("user32.dll")] 
      public static extern IntPtr ReleaseDC(IntPtr hWnd,IntPtr hDC); 
      [DllImport("user32.dll")] 
      public static extern IntPtr GetWindowRect(IntPtr hWnd,ref RECT rect); 
      [DllImport("user32.dll")] 
      public static extern IntPtr GetForegroundWindow();       
    } 
  } 
} 
'@ 
#User Add-Type to import the code 
add-type $code -ReferencedAssemblies 'System.Windows.Forms','System.Drawing' 
#Create the object for the Function 
$capture = New-Object ScreenShotDemo.ScreenCapture 
 
#Take screenshot of the entire screen 
If ($Screen) { 
    Write-Verbose "Taking screenshot of entire desktop" 
    #Save to a file 
    If ($file) { 
        If ($file -eq "") { 
            $file = "$pwd\image.bmp" 
            } 
        Write-Verbose "Creating screen file: $file with imagetype of $imagetype" 
        $capture.CaptureScreenToFile($file,$imagetype) 
        } 
    ElseIf ($print) { 
        $img = $Capture.CaptureScreen() 
        $pd = New-Object System.Drawing.Printing.PrintDocument 
        $pd.Add_PrintPage({$_.Graphics.DrawImage(([System.Drawing.Image]$img), 0, 0)}) 
        $pd.Print() 
        }         
    Else { 
        $capture.CaptureScreen() 
        } 
    } 
#Take screenshot of the active window     
If ($ActiveWindow) { 
    Write-Verbose "Taking screenshot of the active window" 
    #Save to a file 
    If ($file) { 
        If ($file -eq "") { 
            $file = "$pwd\image.bmp" 
            } 
        Write-Verbose "Creating activewindow file: $file with imagetype of $imagetype" 
        $capture.CaptureActiveWindowToFile($file,$imagetype) 
        } 
    ElseIf ($print) { 
        $img = $Capture.CaptureActiveWindow() 
        $pd = New-Object System.Drawing.Printing.PrintDocument 
        $pd.Add_PrintPage({$_.Graphics.DrawImage(([System.Drawing.Image]$img), 0, 0)}) 
        $pd.Print() 
        }         
    Else { 
        $capture.CaptureActiveWindow() 
        }     
    }      
}    



# Declare the variable for getting the test data file location
$myDocumentsFolder = [Environment]::GetFolderPath("MyDocuments")
$file = $myDocumentsFolder + "\Test\test.xlsx"
$sheetName = "Sheet1"

# Create COM object to access the excel application
Try
{
    $objExcel = New-Object -ComObject Excel.Application
    $workbook = $objExcel.Workbooks.Open($file)
    $sheet = $workbook.Worksheets.Item($sheetName)
    $objExcel.Visible=$false
}

Catch [System.Exception]

    {

        $TS = get-date -f MM-dd-yyyy_HH_mm_ss

        "${TS}: Could not load Excel to get test data."

        $message = "FAIL: Could not load Excel to get the test data: "
        $timestamp = get-date -f MM-dd-yyyy_HH:mm:ss

        # Generate HTML report 
        ConvertTo-HTML -Body "$message $timestamp" -Title "Skype Automation Test Report" | Add-Content $myDocumentsFolder\Test\Test_Report_PS.html

        # Open the HTML report
        Invoke-Expression $myDocumentsFolder\Test\Test_Report_PS.html
        Exit

    }

# Create COM object to access the excel document sheet including test  data row & clumn information
$objExcel = New-Object -ComObject Excel.Application
$workbook = $objExcel.Workbooks.Open($file)
$sheet = $workbook.Worksheets.Item(1)
$objExcel.Visible=$false
$rowName,$colName = 22,3
$rowDestName,$colDestName = 22,4
$rowSearch,$colSearch = 22,5

Try
{
    #$Script:username = $sheet.Cells.Item($rowName,$colName).text
    $Script:endusername = $sheet.Cells.Item($rowDestName,$colDestName).text
    #$Script:searchusername = $sheet.Cells.Item($rowSearch,$colSearch).text
     # Close the Excel application object
     $objExcel.quit()
    
}

Catch [System.Exception]

    {

        $TS = get-date -f MM-dd-yyyy_HH_mm_ss

        "${TS}: Could not read Excel to get the test data."
        $message = "FAIL: Could not read Excel to get test data: "
        $timestamp = get-date -f MM-dd-yyyy_HH:mm:ss

        ConvertTo-HTML -Body "$message $timestamp" -Title "Skype Automation Test Report" | Add-Content $myDocumentsFolder\Test\Test_Report_PS.html
        Invoke-Expression $myDocumentsFolder\Test\Test_Report_PS.htm
        # Close the Excel application object
        $objExcel.quit()
        Exit
    
    }

# Load the Lync 2013 SDK to work with API's 
try
{
$assemblyPath = 'C:\Program Files (x86)\Microsoft Office 2013\LyncSDK\Assemblies\Desktop\Microsoft.Lync.Model.DLL'
Import-Module $assemblyPath

    
 $IMType = 2
 $Client = [Microsoft.Lync.Model.LyncClient]::GetClient()
 $conv = $client.ConversationManager.AddConversation()
 $getuser = $client.ContactManager.GetContactByUri($Script:endusername)
 $result = $getuser.GetContactInformation("Availability")
 $conv.AddParticipant($getuser)
 $m = $conv.Modalities[$IMType]

 if($result -eq 0 -or $result -eq 18500)
 {
     write-host "Skype contact is not available."
      
        "Search: Contact is not available in Skype or LDAP."
        
        $message = "FAIL: Search: Contact is not available in Skype or LDAP: "
        $timestamp = get-date -f MM-dd-yyyy_HH:mm:ss

        ConvertTo-HTML -Body "$message $timestamp" -Title "Skype Automation Test Report" | Add-Content $myDocumentsFolder\Test\Test_Report_PS.html
 }
 else
 {
   
    $null=$m.BeginConnect($null,$null)
    Start-Sleep -s 15

     
     # Calling the Take-ScreenShot functiona for screen capture and saved in the respective location for test evidence
     $timestamp = get-date -f MM-dd-yyyy_HH_mm_ss
    Take-ScreenShot -screen -file "$myDocumentsFolder\Test\Skype_Voice_Call_Triggered_$timestamp.jpeg" -imagetype jpeg 

        "Voice call successfully triggered."
        
        $message = "PASS: Voice call successfully triggered: "
        $timestamp = get-date -f MM-dd-yyyy_HH:mm:ss

        ConvertTo-HTML -Body "$message $timestamp" -Title "Skype Automation Test Report" | Add-Content $myDocumentsFolder\Test\Test_Report_PS.html
    $m.Participant.BeginSetMute("true",{},"true")
    Start-Sleep -s 5

    
     # Calling the Take-ScreenShot functiona for screen capture and saved in the respective location for test evidence
      $timestamp = get-date -f MM-dd-yyyy_HH_mm_ss
    Take-ScreenShot -screen -file "$myDocumentsFolder\Test\Skype_Mute_Success_$timestamp.jpeg" -imagetype jpeg 

     "Mute option enabled."
     $message = "PASS: Mute option enabled: "
        $timestamp = get-date -f MM-dd-yyyy_HH:mm:ss

        ConvertTo-HTML -Body "$message $timestamp" -Title "Skype Automation Test Report" | Add-Content $myDocumentsFolder\Test\Test_Report_PS.html
    $m.BeginSetProperty("AVModalityAudioCaptureMute",$false,$UnMuteOccured, $null)   
    Start-Sleep -s 5

     # Calling the Take-ScreenShot functiona for screen capture and saved in the respective location for test evidence
     $timestamp = get-date -f MM-dd-yyyy_HH_mm_ss
    Take-ScreenShot -screen -file "$myDocumentsFolder\Test\Skype_Unmute_Success_$timestamp.jpeg" -imagetype jpeg 

    "Mute option disabled."
     $message = "PASS: Mute option disabled: "
        $timestamp = get-date -f MM-dd-yyyy_HH:mm:ss

        ConvertTo-HTML -Body "$message $timestamp" -Title "Skype Automation Test Report" | Add-Content $myDocumentsFolder\Test\Test_Report_PS.html
    
if ($m.state -eq "Connected")
 {
    write-host "Successfully Connected Voice Call"
    # Video call
    $m.VideoChannel.BeginStart({},0) 
    write-host "Successfully Connected video Call"
    Start-Sleep -s 15
     # Calling the Take-ScreenShot functiona for screen capture and saved in the respective location for test evidence
    $timestamp = get-date -f MM-dd-yyyy_HH_mm_ss
    Take-ScreenShot -screen -file "$myDocumentsFolder\Test\Skype_Video_Call_Success_$timestamp.jpeg" -imagetype jpeg 
             
        $message = "PASS: Video call successfully connected: "
        $timestamp = get-date -f MM-dd-yyyy_HH:mm:ss

        ConvertTo-HTML -Body "$message $timestamp" -Title "Skype Automation Test Report" | Add-Content $myDocumentsFolder\Test\Test_Report_PS.html
    # share my desktop
    $conv.Modalities['ApplicationSharing'].BeginShareDesktop({}, 0)
    
    Start-Sleep -s 10

     # Calling the Take-ScreenShot functiona for screen capture and saved in the respective location for test evidence
     $timestamp = get-date -f MM-dd-yyyy_HH_mm_ss
    Take-ScreenShot -screen -file "$myDocumentsFolder\Test\Skype_Desktop_Sharing_Connect_Success_$timestamp.jpeg" -imagetype jpeg 

    write-host "Successfully Shared the Desktop"
     $message = "PASS: Desktop sharing successfully connected: "
        $timestamp = get-date -f MM-dd-yyyy_HH:mm:ss

        ConvertTo-HTML -Body "$message $timestamp" -Title "Skype Automation Test Report" | Add-Content $myDocumentsFolder\Test\Test_Report_PS.html
    $conv.Modalities['ApplicationSharing'].BeginDisconnect([Microsoft.Lync.Model.Conversation.ModalityDisconnectReason]::None, {}, 0)   
    Start-Sleep -s 5

     # Calling the Take-ScreenShot functiona for screen capture and saved in the respective location for test evidence
     $timestamp = get-date -f MM-dd-yyyy_HH_mm_ss
    Take-ScreenShot -screen -file "$myDocumentsFolder\Test\Skype_Desktop_Sharing_Disconnect_Success_$timestamp.jpeg" -imagetype jpeg 

    $message = "PASS: Desktop sharing successfully disconnected: "
        $timestamp = get-date -f MM-dd-yyyy_HH:mm:ss

        ConvertTo-HTML -Body "$message $timestamp" -Title "Skype Automation Test Report" | Add-Content $myDocumentsFolder\Test\Test_Report_PS.html
    
}
else
{
    write-host "Destination user is not accepting Voice Call"
    $message = "FAIL: Destination not accepting Voice Call: "
        $timestamp = get-date -f MM-dd-yyyy_HH:mm:ss

        ConvertTo-HTML -Body "$message $timestamp" -Title "Skype Automation Test Report" | Add-Content $myDocumentsFolder\Test\Test_Report_PS.html
}

write-host "Ending Call"
#end the call
$m.Conversation.End()
}
}
 catch
    {
        $_.Exception 
        $TS = get-date -f MM-dd-yyyy_HH_mm_ss

        "${TS}: Exception in voice call triggering process."
        $message = "FAIL: Problem in voice call triggering process: "
        $timestamp = get-date -f MM-dd-yyyy_HH:mm:ss

        ConvertTo-HTML -Body "$message $timestamp" -Title "Skype Automation Test Report" | Add-Content $myDocumentsFolder\Test\Test_Report_PS.html

        
        $m.Conversation.End()
        Invoke-Expression $myDocumentsFolder\Test\Test_Report_PS.html
        Exit
    }

