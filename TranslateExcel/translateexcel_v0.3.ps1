param(
[Parameter(Mandatory=$True,Position=1)]
   [string]$excelworkbook,
[Parameter(Mandatory=$True,Position=2)]
   [string]$excelworksheet,
[string]$fromLanguageCode = "auto",
[String]$toLanguageCode = "en"
)

$excelworkbook = "D:\wbooktests.xlsx"
$excelworksheet = "Sheet1"
echo "Only stuff already saved in '$excelworkbook' file, in '$excelworksheet' sheet will be translated. Save your file first if there are unsaved changes"
#########################################################################
#repeatable function definition
function translate-this ()
{
#by default it converts "Saurabh loco poco" to english
#refer translate.google.com for google language codes.
param (
    [string]$textToTranslate = "Saurabh loco poco",
    [string]$sourceLanguageCode = "auto",
    [string]$destinationLanguageCode = "en",
    [System.__ComObject] $ie 
    )
#echo "Attempting to translate : '$textToTranslate'"
#$ie = New-Object -ComObject InternetExplorer.Application
$translatedText = "*Couldn't translate*"

[System.Reflection.Assembly]::LoadWithPartialName("System.web") | out-null
$urlEncodedTextToTranslate = [System.Web.HttpUtility]::UrlEncode($textToTranslate) 

$ie.Navigate("https://translate.google.com/#$sourceLanguageCode/$destinationLanguageCode/$urlEncodedTextToTranslate")
while ($ie.Busy -eq $true) {
    Start-Sleep -Milliseconds 50
}

$doc = $ie.Document

try {
    $detectedLanguage = $doc.getElementById("result_box")
    $resultBox = $doc.getElementById("result_box")
    $translatedText= $resultBox.innerText
    
} catch {
           echo "*translationfailed*"
           return $translatedText
        }
return $translatedText
}


$workingdirectory = Split-Path $excelworkbook
$filename = $excelworkbook.Split("\").item( $excelworkbook.Split("\").count - 1)

$ie = New-Object -ComObject InternetExplorer.Application
try{
$xl = New-Object -comobject Excel.Application
$xl.visible = $true
$xl.DisplayAlerts = $false
} catch {
            Write-Error "le bhai: couldn't even create excel com-object"
        }

$wbook = $xl.Workbooks.open($excelworkbook)
$wbook.SaveAs($excelworkbook + "_translated.xlsx")

foreach ($rangecell in $xl.workbooks.item($wbook.name).worksheets.item($excelworksheet).usedrange)
{
    $dontneedtotranslate = $false
    $rangecell.cells.value  -match "[1234567890]*" | out-null
    $dontneedtotranslate = $Matches -eq $rangecell.cells.value
    if($dontneedtotranslate ) {
        $dontneedtotranslate
        continue
    }
    
    
    $trans =  translate-this -textToTranslate $rangecell.cells.value() -sourceLanguageCode $fromLanguageCode -destinationLanguageCode $toLanguageCode -ie $ie
    $xl.workbooks.item($wbook.name).worksheets.item($excelworksheet).range($rangecell.address()).cells.value2 = $trans
    echo "last translated: $tran" 
}

$wbook.save()


##Quit IE and Excel. Not graceful at all, full blood and gore
$ie.Quit()
##please refer: https://technet.microsoft.com/en-us/library/ff730962.aspx ##code to get rid of the excel instance, a simple $xl.quit does not kill the background process. --v-nagaur
$xl.Quit()
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($xl)
Remove-Variable xl

"Finished Execution"

#if there is anything you do not understand in this code please visit: "https://www.google.co.in/search?q=earth+day+quiz&oi=ddle&ct=ddle-hpp&hl=en"