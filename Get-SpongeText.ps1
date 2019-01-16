param([string]$folderPath="C:\users\username\desktop\test*")
#$folderPath = "C:\Users\ethomas\Desktop\test\1*" # multi-folders: "C:\fso1*", "C:\fso2*"
$fileType = "*.doc" # *.doc will take all .doc* files

$word = New-Object -ComObject Word.Application
$word.Visible = $false
# problem. It only finds first occurence. 
Function findAndReplace($Text, $Find, $ReplaceWith) {
    $matchCase = $true
    $matchWholeWord = $false
    $matchWildcards = $false
    $matchSoundsLike = $false
    $matchAllWordForms = $false
    $forward = $true
    $findWrap = [Microsoft.Office.Interop.Word.WdReplace]::wdReplaceAll
    $format = $false
    $replace = [Microsoft.Office.Interop.Word.WdFindWrap]::wdFindContinue

    $Text.Execute($Find, $matchCase, $matchWholeWord, $matchWildCards, ` 
                  $matchSoundsLike, $matchAllWordForms, $forward, $findWrap, `  
                  $format, $ReplaceWith, $replace) > $null
}

Function findAndReplaceWholeDoc($Document, $Find, $ReplaceWith) {
    $findReplace = $Document.ActiveWindow.Selection.Find
    findAndReplace -Text $findReplace -Find $Find -ReplaceWith $ReplaceWith
    ForEach ($section in $Document.Sections) {
        ForEach ($header in $section.Headers) {
            $findReplace = $header.Range.Find
            findAndReplace -Text $findReplace -Find $Find -ReplaceWith $ReplaceWith
            $header.Shapes | ForEach-Object {
                if ($_.Type -eq [Microsoft.Office.Core.msoShapeType]::msoTextBox) {
                    $findReplace = $_.TextFrame.TextRange.Find
                    findAndReplace -Text $findReplace -Find $Find -ReplaceWith $ReplaceWith
                }
            }
        }
        ForEach ($footer in $section.Footers) {
            $findReplace = $footer.Range.Find
            findAndReplace -Text $findReplace -Find $Find -ReplaceWith $ReplaceWith
        }
    }
}
# 97 through 122 is lowercase
# 65 to 90 is uppercase. 
# 97..122 | ForEach-Object {[Char]$PSItem}
# 65..90 | ForEach-Object {[Char]$PSItem}

Function processDoc {
    $doc = $word.Documents.Open($_.FullName)
    #
    # Replace lower to upper
    
    for ($increment1 = 97; $increment1 -lt 123; $increment1++)
    {
        $letter =  [char]($increment1)
        $random = Get-Random
        if(($random % 3 ) -eq 0){
            $upperCaseLetter = [char]([char]($increment1) - 32)
            Write-Output "Replace $letter to $upperCaseLetter"
            findAndReplaceWholeDoc -Document $doc -Find "$letter" -ReplaceWith "$upperCaseLetter"
        }
    } 
     
    
    #Replace upper to lower
    for ($increment2 = 65; $increment2 -lt 91; $increment2++)
    {
        $letter =  [char]($increment2)
        $random = Get-Random
        if(($random % 3 ) -eq 0){
            $lowercaseLetter = [char]([char]($increment2 + 32))
            #$lowercaseLetter = ([char]($test))
            Write-Output "Replace $letter to $lowercaseLetter"
            findAndReplaceWholeDoc -Document $doc -Find "$letter" -ReplaceWith "$lowercaseLetter"
        }
    } 

    
    
    $doc.Close([ref]$true)
}

$sw = [Diagnostics.Stopwatch]::StartNew()
$count = 0
Get-ChildItem -Path $folderPath -Recurse -Filter $fileType | ForEach-Object { 
  Write-Host "Processing \`"$($_.Name)\`"..."
  processDoc
  $count++
}
$sw.Stop()
$elapsed = $sw.Elapsed.toString()
Write-Host "`nDone. $count files processed in $elapsed" 

$word.Quit()
$word = $null
[gc]::collect() 
[gc]::WaitForPendingFinalizers()
