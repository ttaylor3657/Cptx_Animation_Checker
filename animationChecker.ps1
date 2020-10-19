# *************************************************************************
#
# Script Name: Animation Checker
# Version: 0.1
# Author: Jeff Taylor
# Date: 5/14/2020
#
# Description: Run script and it will open a GUI. Select a lesson package that cotains
# HTML5 animations. A report will be generated with a list of animations that do not inlcude
# the proper javascript in the file, and use the hosted version. These can then be corrected
# by the artists and/or programmer.
#
#   WARNING: Nested zip files are no supported!
#   
# *************************************************************************

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.IO
Add-Type -AssemblyName System.IO.Compression.FileSystem
Add-Type -AssemblyName Microsoft.VisualBasic

$script:j = 0

#region Functions

#This will open and return a file stream from a zipped file, requires zip file path and internal location
function Read-FileInZip($ZipFilePath, $FilePathInZip) {
    try {
        if (![System.IO.File]::Exists($ZipFilePath)) {
            throw "Zip file $ZipFilePath not found."
        }

        $Zip = [System.IO.Compression.ZipFile]::OpenRead($ZipFilePath)
        $ZipEntries = [array]($Zip.Entries | where-object {
                return $_.FullName -eq $FilePathInZip
            });
        if (!$ZipEntries -or $ZipEntries.Length -lt 1) {
            throw "File $FilePathInZip couldn't be found in zip $ZipFilePath."
        }
        if (!$ZipEntries -or $ZipEntries.Length -gt 1) {
            throw "More than one file $FilePathInZip found in zip $ZipFilePath."
        }

        $ZipStream = $ZipEntries[0].Open()

        $Reader = [System.IO.StreamReader]::new($ZipStream)
        return $Reader.ReadToEnd()
    }
    finally {
        if ($Reader) { $Reader.Dispose() }
        if ($Zip) { $Zip.Dispose() }
    }
}


#This function will process a list of zip files, and call the Test-HTML  
function get-MultipleLesson($fileList){
    $zippedAni = New-Object System.Collections.Generic.List[System.Object]

    $zFiles = $fileList -match "zip"

#Processing Zip files
    $OutputTextBox.AppendText("`r`nProcessing " + $zFiles.count +" lesson package files. Please wait...")
    ForEach ($zipFile in $zFiles) {

        $pBarCount++
        $PBarPercent = ($pBarCount/$zFiles.Count)*100
        $ProgressBar1.Value = $PBarPercent

        if ( { $_.name -like '*.zip' }) {
             $zip = [System.IO.Compression.ZipFile]::OpenRead($zipFile)

             $ZipEntries = [array]($zip.Entries | where-object {
                return $_.FullName -match '(?!\S*xlp)(?!\S*index)wor.*\.html?'
            });
            
            for ($i = 0; $i -lt $ZipEntries.Count; $i++) {
                $zItem = $zipFile + "|" + $ZipEntries[$i]
                $zippedAni.add($zItem)
            }

            $zip.Dispose()
        }
    }
    
    $OutputTextBox.AppendText("`r`nAnalyzing " + $zippedAni.count +" animations. Please wait...")
    $pBarCount = 0
    $PBarPercent = 0
    $ProgressBar1.Value = 0

    for ($i = 0; $i -lt $zippedAni.Count; $i++) {

        $pBarCount++
        $PBarPercent = ($pBarCount/$zippedAni.Count) * 100
        $ProgressBar1.Value = $PBarPercent

        $currentListItem = $zippedAni[$i].split("|")
        $myHtml = Read-FileInZip $currentListItem[0]  $currentListItem[1]
        
        if($currentListItem[0] -ne $tempZipName){
            $tempZipName = $currentListItem[0]
            $zipName= "`r`nAnalyzing " + $currentListItem[0]
            $OutputTextBox.AppendText($zipName)
        }
        
        test-HTML $myHtml $currentListItem[1] $j

    }
    switch ($script:j) {
        0       {$OutputTextBox.AppendText("`r`nAnalysis complete. No animations need to be corrected.")}
        1       {$OutputTextBox.AppendText("`r`nAnalysis complete. $script:j animation needs to be corrected.")}
        default {$OutputTextBox.AppendText("`r`nAnalysis complete. $script:j animations need to be corrected.")}
    }
}

#Find the Captivate preview temp html folder, and call test-HTML for each HTML5 Animation
function get-HtmlPreview($path){
    $subValue = $path.indexof("HTML5") 
    $searchPath = $path.Substring($subValue)
    $subValue = $searchPath.indexof("/")
    $searchPath = $searchPath.Substring(0, $subValue)
 
    $lessonPath = Get-ChildItem -Path $env:TEMP -recurse | Where-Object {$_.name -like $searchPath}

    $lessonHtml = Get-Content ($lessonpath.FullName + '\index.html')
    $lessonTitle = [regex]::Match($lessonHtml, '<title>(?<title>.*)<\/title>')
    if($lessonTitle.Success){
        $OutputTextBox.AppendText("`r`nProcessing Captivate HTML5 Preview for " + $lessonTitle.Groups['title'].value + ". Please wait...")
    } else {
        $OutputTextBox.AppendText("`r`nProcessing Captivate HTML5 Preview. Please wait...")
    }

    $webObjects = Get-ChildItem -Path ($lessonPath.FullName + "\wor") -Recurse -Filter "*.html" | Where-Object {$_.FullName -match '(?!\S*xlp)(?!\S*index)wor.*\.html?'}

    foreach ($file in $webObjects){
        $htmlData = Get-Content $file.FullName

        test-html $htmlData $file.name $j
    }

    switch ($script:j) {
        0       {$OutputTextBox.AppendText("`r`nAnalysis complete. No animations need to be corrected.")}
        1       {$OutputTextBox.AppendText("`r`nAnalysis complete. $script:j animation needs to be corrected.")}
        default {$OutputTextBox.AppendText("`r`nAnalysis complete. $script:j animations need to be corrected.")}
    }
}

#This checks an HTML file for a
function test-HTML($htmlData, $htmlName, $j) {

    if($htmlData -match ("https://code.createjs.com/")){
        $OutputTextBox.AppendText("`r`n     $htmlName will not work in the learning center.")
        $script:j++
    }
}

#Called when analyze button clicked, checks path is valid
function test-Animation {
    $script:j = 0

    if($pathTextBox.text -like "*.html"){
        get-HtmlPreview($pathTextBox.text)
    } else {
        $itemsToBeConverted = $pathTextBox.text.split(";")
        
        if($itemsToBeConverted.Count -le 0){
            $OutputTextBox.AppendText("`r`nSource path not valid.")
            break
        } 
        
        if ($itemsToBeConverted.Count -gt 1){
            get-MultipleLesson $itemsToBeConverted

        } elseif ($itemsToBeConverted.Count -eq 1) {

            if($itemsToBeConverted -like "*.zip"){
                get-MultipleLesson $itemsToBeConverted

            } else {
                $matchedFiles = Get-ChildItem $itemsToBeConverted -Include "*.zip" -Recurse
                $fileList = @()

                foreach($file in $matchedFiles){
                    $fileList += $file.FullName
                }

                get-MultipleLesson $fileList
            }
        }
    }
 }
 
function get-SourceFolderDialog {
    $FolderBrowser = New-Object System.Windows.Forms.FolderBrowserDialog 
    $FolderBrowser.ShowDialog() | Out-Null
    $sourceFolder = $FolderBrowser.SelectedPath
    $pathTextBox.text = $sourceFolder

    if ($pathTextBox.text.length -gt 0) {
        $analyzeButton.enabled = $true
    }
}

function get-SourceFileDialog {
    $FileBrowser = New-Object System.Windows.Forms.OpenFileDialog -Property @{ 
        InitialDirectory = [Environment]::GetFolderPath('Desktop') 
        Filter = 'Files (*.zip;)|*.zip;|Zip (*.zip)|*.zip'
        Multiselect = $true
    }

    $FileBrowser.ShowDialog() | Out-Null
    foreach ($file in $FileBrowser.FileNames) {        
        $sourceFile += $file + ";"
    }
    if($null -ne $sourceFile){
    $sourceFile = $sourceFile.trimend(";")
    }
    $pathTextBox.text = $sourceFile

    if ($pathTextBox.text.length -gt 0) {
        $analyzeButton.enabled = $true
    }
}

function get-SourceHtmlDialog {
    $title = 'HTML5 Animation Test Path'
    $msg = "Enter the full URL from the Captivate HTML5 preview broswer window."
    $pathTextBox.text = [Microsoft.VisualBasic.Interaction]::InputBox($msg, $title)

    if ($pathTextBox.text.length -gt 0) {
        $analyzeButton.enabled = $true
    }
}

function clear-log {
    $OutputTextBox.Text = ""
}
function copy-log {
    Set-Clipboard -Value $OutputTextBox.Text
}

function save-log {
    $saveFileDialog = New-Object System.Windows.Forms.SaveFileDialog -Property @{ 
        InitialDirectory = [Environment]::GetFolderPath('Desktop') 
        Filter           = 'Files (*.txt;)|*.txt;|Txt (*.txt)|*.txt'
    }

    $saveFileDialog.ShowDialog() | Out-Null

    if ($saveFileDialog.FileName -ne "") {
        $text = $OutputTextBox.Text
        Set-Content -Path $saveFileDialog.FileName -Value $text
    }
}

#endregion Functions

#region GUI
[System.Windows.Forms.Application]::EnableVisualStyles()

$AnimationChecker                  = New-Object system.Windows.Forms.Form
$AnimationChecker.ClientSize       = '400,509'
$AnimationChecker.text             = "Animation Checker"
#$AnimationChecker.BackColor        = "#4a4a4a"
$AnimationChecker.TopMost          = $false
$AnimationChecker.MaximizeBox      = $false
$AnimationChecker.FormBorderStyle  = "Fixed3D"

$inputGroupBox                     = New-Object system.Windows.Forms.Groupbox
$inputGroupBox.height              = 120
$inputGroupBox.width               = 378
$inputGroupBox.location            = New-Object System.Drawing.Point(9,75)

$outputGroupBox                    = New-Object system.Windows.Forms.Groupbox
$outputGroupBox.height             = 45
$outputGroupBox.width              = 378
$outputGroupBox.location           = New-Object System.Drawing.Point(9,200)

$fileBroswerButton                 = New-Object system.Windows.Forms.Button
#$fileBroswerButton.BackColor        = "#4a4a4a"
$fileBroswerButton.text            = "Select package"
$fileBroswerButton.width           = 173
$fileBroswerButton.height          = 30
$fileBroswerButton.location        = New-Object System.Drawing.Point(14,18)
$fileBroswerButton.Font            = 'Microsoft Sans Serif,10'
#$fileBroswerButton.ForeColor        = "#ffffff"

$folderBroswerButton               = New-Object system.Windows.Forms.Button
#$folderBroswerButton.BackColor     = "#4a4a4a"
$folderBroswerButton.text          = "Select folder"
$folderBroswerButton.width         = 171
$folderBroswerButton.height        = 30
$folderBroswerButton.location      = New-Object System.Drawing.Point(196,18)
$folderBroswerButton.Font          = 'Microsoft Sans Serif,10'
#$folderBroswerButton.ForeColor     = "#ffffff"

$html5BroswerButton                = New-Object system.Windows.Forms.Button
#$html5BroswerButton.BackColor     = "#4a4a4a"
$html5BroswerButton.text           = "Select HTML5 Preview"
$html5BroswerButton.width          = 171
$html5BroswerButton.height         = 30
$html5BroswerButton.location       = New-Object System.Drawing.Point(106,52)
$html5BroswerButton.Font           = 'Microsoft Sans Serif,10'
#$html5BroswerButton.ForeColor      = "#ffffff"

$pathTextBox                       = New-Object system.Windows.Forms.TextBox
$pathTextBox.multiline             = $false
#$pathTextBox.BackColor             = "#2a2a2a"
$pathTextBox.width                 = 353
$pathTextBox.height                = 30
$pathTextBox.location              = New-Object System.Drawing.Point(14,87)
$pathTextBox.Font                  = 'Microsoft Sans Serif,10'
#$pathTextBox.ForeColor             = "#ffffff"
$pathTextBox.ReadOnly              = $true

$analyzeButton                     = New-Object system.Windows.Forms.Button
#$analyzeButton.BackColor           = "#4a4a4a"
$analyzeButton.text                = "Check animation files"
$analyzeButton.width               = 353
$analyzeButton.height              = 30
$analyzeButton.enabled             = $false
$analyzeButton.location            = New-Object System.Drawing.Point(15,10)
$analyzeButton.Font                = 'Microsoft Sans Serif,10'
#$analyzeButton.ForeColor           = "#ffffff"

$Groupbox1                         = New-Object system.Windows.Forms.Groupbox
$Groupbox1.height                  = 244
$Groupbox1.width                   = 378
$Groupbox1.text                    = "Output"
$Groupbox1.location                = New-Object System.Drawing.Point(10,255)

$ProgressBar1                      = New-Object system.Windows.Forms.ProgressBar
#$ProgressBar1.BackColor            = "#7ed321"
$ProgressBar1.width                = 356
$ProgressBar1.height               = 17
$ProgressBar1.location             = New-Object System.Drawing.Point(10,20) 
$ProgressBar1.Value                = 0
$ProgressBar1.Style                = "Continuous"

$OutputTextBox                     = New-Object system.Windows.Forms.TextBox
$OutputTextBox.multiline           = $true
$OutputTextBox.ScrollBars          = "Vertical"
#$OutputTextBox.BackColor           = "#2a2a2a"
$OutputTextBox.width               = 356
$OutputTextBox.height              = 150
$OutputTextBox.location            = New-Object System.Drawing.Point(10,42)
$OutputTextBox.Font                = 'Microsoft Sans Serif,10'
#$OutputTextBox.ForeColor           = "#ffffff"
$OutputTextBox.ReadOnly            = $true

$copyLogButton                     = New-Object system.Windows.Forms.Button
$copyLogButton.text                = "Copy results"
$copyLogButton.width               = 100
$copyLogButton.height              = 30
$copyLogButton.location            = New-Object System.Drawing.Point(10,202)
$copyLogButton.Font                = 'Microsoft Sans Serif,10'

$saveLogButton                     = New-Object system.Windows.Forms.Button
$saveLogButton.text                = "Save results"
$saveLogButton.width               = 100
$saveLogButton.height              = 30
$saveLogButton.location            = New-Object System.Drawing.Point(140,202)
$saveLogButton.Font                = 'Microsoft Sans Serif,10'

$clearLogButton                    = New-Object system.Windows.Forms.Button
$clearLogButton.text               = "Clear log"
$clearLogButton.width              = 100
$clearLogButton.height             = 30
$clearLogButton.location           = New-Object System.Drawing.Point(265,202)
$clearLogButton.Font               = 'Microsoft Sans Serif,10'

$Label1                            = New-Object system.Windows.Forms.Label
$Label1.text                       = "This tool checks animations in published Captivate lesson packages (zip) or in local Captivate previews (url). Previews in Captivate must be generated using the HTML5 in Broswer option."
$Label1.AutoSize                   = $false
$Label1.width                      = 380
$Label1.height                     = 100
$Label1.location                   = New-Object System.Drawing.Point(10,12)
$Label1.Font                       = 'Microsoft Sans Serif,10'
#$Label1.ForeColor                  = "#ffffff"

$AnimationChecker.controls.AddRange(@($inputGroupBox,$outputGroupBox,$Groupbox1,$Label1))
$inputGroupBox.controls.AddRange(@($fileBroswerButton,$pathTextBox,$folderBroswerButton,$html5BroswerButton))
$outputGroupBox.controls.AddRange(@($analyzeButton))
$Groupbox1.controls.AddRange(@($ProgressBar1,$OutputTextBox,$copyLogButton,$saveLogButton,$clearLogButton))

$fileBroswerButton.Add_Click({ get-SourceFileDialog })
$folderBroswerButton.Add_Click({ get-SourceFolderDialog })
$html5BroswerButton.Add_Click({ get-SourceHtmlDialog })
$analyzeButton.Add_Click({ test-Animation })
$copyLogButton.Add_Click({ copy-log })
$saveLogButton.Add_Click({ save-log })
$clearLogButton.Add_Click({ clear-log })


#endregion GUI

#Call the form
[void]$AnimationChecker.ShowDialog()