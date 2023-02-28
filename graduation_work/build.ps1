param (
  [Parameter(Mandatory = $true)][string]$reference,
  [Parameter(Mandatory = $true)][string[]]$md,  
  [string]$template,
  [string]$docx,
  [switch]$embedfonts,
  [switch]$counters
)

if ([string]::IsNullOrEmpty($docx)) {
  Write-error "Необходимо определить результирующий документ docx"
  exit 113
}

$finalRecommendations = "РЕКОМЕНДАЦИИ:"
$pandoc = "pandoc"

$reference = [System.IO.Path]::GetFullPath($reference)

$md = $md | % { [System.IO.Path]::GetFullPath($_ ) }

$is_docx_temporary = $False

$docx = [System.IO.Path]::GetFullPath($docx)

$tempdocx = [System.IO.Path]::GetTempFilename() + ".docx"

Write-Host "Запуск Pandoc..."
& $pandoc $md -o $tempdocx --lua-filter ./linebreaks.lua --filter pandoc-crossref --citeproc --reference-doc $reference

if ($LASTEXITCODE -ne 0) {
  Write-error "Во время работы pandoc произошла ошибка"
  exit 111
}

$word = New-Object -ComObject Word.Application
$curdir = Split-Path $MyInvocation.MyCommand.Path
Set-Location -Path $curdir

$word.ScreenUpdating = $False

$template = [System.IO.Path]::GetFullPath($template)
$doc = $word.Documents.Open($template)
$doc.Activate()
$selection = $word.Selection

# Save under a new name as soon as possible to prevent auto-saving
# (and polluting) the template
$doc.SaveAs([ref]$docx)
if (-not $?) {
  $doc.Close()
  $word.Quit()
  exit 112
}

# Disable grammar checking (it takes time and spews out error messages)
$doc.GrammarChecked = $True

Write-Host "Вставка основного текста..."
if ($selection.Find.Execute("%MAINTEXT%^13", $True, $True, $False, $False, $False, $True, `
      [Microsoft.Office.Interop.Word.wdFindWrap]::wdFindContinue, $False, "", `
      [Microsoft.Office.Interop.Word.wdReplace]::wdReplaceNone)) {
  $Selection.InsertFile($tempdocx)

  # Check if there is anything after the main text
  $selection.WholeStory()
  $totalend = $Selection.Range.End

  # If there is nothing after the main text, remove the extra CR which
  # mystically appears out of nowhere in that case
  if ($end -ge ($totalend - 1)) {
    $selection.Collapse([Microsoft.Office.Interop.Word.wdCollapseDirection]::wdCollapseEnd)  | out-null
    $selection.MoveLeft([Microsoft.Office.Interop.Word.wdUnits]::wdCharacter, 1, `
        [Microsoft.Office.Interop.Word.wdMovementType]::wdExtend)  | out-null
    $selection.Delete() | out-null
  }
}

# Удаление всех гиперссылок
$hyperlinks = @($doc.Hyperlinks) 
$hyperlinks | ForEach {
  $_.Delete()
}

Write-Host "Поиск стилей..."
foreach ($style in $doc.Styles) {
  switch ($style.NameLocal) {
    'TableStyleContributors' { $TableStyleContributors = $style; break }
    'TableStyleAbbreviations' { $TableStyleAbbreviations = $style; break }
    'TableStyleGost' { $TableStyleGost = $style; break }
    'TableStyleGostNoHeader' { $TableStyleGostNoHeader = $style; break }
    'UnnumberedHeading1' { $UnnumberedHeading1 = $style; break }
    'UnnumberedHeading1NoTOC' { $UnnumberedHeading1NoTOC = $style; break }
    'UnnumberedHeading2' { $UnnumberedHeading2 = $style; break }
  }
}

$BodyText = [Microsoft.Office.Interop.Word.wdBuiltinStyle]::wdStyleBodyText
$Heading1 = [Microsoft.Office.Interop.Word.wdBuiltinStyle]::wdStyleHeading1
$Heading2 = [Microsoft.Office.Interop.Word.wdBuiltinStyle]::wdStyleHeading2
$Heading3 = [Microsoft.Office.Interop.Word.wdBuiltinStyle]::wdStyleHeading3

$bullets = [char]0x2014, [char]0xB0, [char]0x2014, [char]0xB0
$numberposition = 1.25 # Отступ списка
$textposition = 0
$tabposition = 0, 1.75, 3, 3.5
$format_nested = "%1)", "%1.%2)", "%1.%2.%3)", "%1.%2.%3.%4)"
$format_headers = "%1", "%1.%2", "%1.%2.%3", "%1.%2.%3.%4"
$format_single = "%1)", "%2)", "%3)", "%4)"

Write-Host "Шаблоны списков..."

foreach ($tt in $doc.ListTemplates) {
  for ($il = 1; $il -le $tt.ListLevels.Count -and $il -le 4; $il++) {
    $level = $tt.ListLevels.Item($il)
    $bullet = ($level.NumberStyle -eq [Microsoft.Office.Interop.Word.wdListNumberStyle]::wdListNumberStyleBullet)
    $arabic = ($level.NumberStyle -eq [Microsoft.Office.Interop.Word.wdListNumberStyle]::wdListNumberStyleArabic)
    $roman = ($level.NumberStyle -eq [Microsoft.Office.Interop.Word.wdListNumberStyle]::wdListNumberStyleLowercaseRoman)
    
    if ($bullet) {
      if ($level.NumberFormat -ne " ") {
        $level.NumberFormat = $bullets[$il - 1] + ""
      }
      $level.NumberPosition = $word.CentimetersToPoints($numberposition)
      $level.Alignment = [Microsoft.Office.Interop.Word.wdListLevelAlignment]::wdListLevelAlignLeft
      $level.TextPosition = $word.CentimetersToPoints($textposition)
      $level.TabPosition = $word.CentimetersToPoints($tabposition[$il - 1])
      $level.ResetOnHigher = $il - 1
      $level.StartAt = 1
      $level.Font.Size = 14
      $level.Font.Name = "Times New Roman"
      if ($il % 2 -eq 0) {
        $level.Font.Position = -4
      }
      $level.LinkedStyle = ""
      $level.TrailingCharacter = [Microsoft.Office.Interop.Word.wdTrailingCharacter]::wdTrailingTab
    }
    
    if (($arabic -and ($level.NumberFormat -ne $format_headers[$il - 1])) -or $roman) {
      if ($level.NumberFormat -ne " " ) {
        if ($arabic) {
          $level.NumberFormat = $format_nested[$il - 1]
        }
        if ($roman) {
          $level.NumberStyle = [Microsoft.Office.Interop.Word.wdListNumberStyle]::wdListNumberStyleArabic;
          $level.NumberFormat = $format_single[$il - 1]
        }
      }
      $level.NumberPosition = $word.CentimetersToPoints($numberposition)
      $level.Alignment = [Microsoft.Office.Interop.Word.wdListLevelAlignment]::wdListLevelAlignLeft
      $level.TextPosition = $word.CentimetersToPoints($textposition)
      $level.TabPosition = $word.CentimetersToPoints($tabposition[$il - 1])
      $level.ResetOnHigher = $il - 1
      $level.StartAt = 1
      $level.Font.Size = 14
      $level.Font.Name = "Times New Roman"
      $level.LinkedStyle = ""
      $level.TrailingCharacter = [Microsoft.Office.Interop.Word.wdTrailingCharacter]::wdTrailingTab
    }
  }
}

$ntables = 0

Write-Host "Обработка таблиц..."
$user_tables = $doc.Tables

for ($t = 1; $t -le $user_tables.Count; $t++) {

  $table = $user_tables.Item($t)
  # Если таблица сделана для формулы
  if ([string]::IsNullOrEmpty($table.Title) `
      -and ($table.Columns.Count -eq 2) `
      -and ($table.Cell(1, 1).Range.OMaths.Count -eq 1) `
      -and ($table.Cell(1, 2).Range.OMaths.Count -eq 1)) {    

    # There can be multiple equations (rows) in one table
    foreach ($row in $table.Rows) {
      # After removing the equation, the text contents remains
      if ($row.Cells.Item(2).Range.OMaths.Count -ne 0) {
        $row.Cells.Item(2).Range.OMaths.Item(1).Remove()
      }

      foreach ($border in $row.Cells.Borders) {
        $border.Visible = $False
      }
      $row.Cells.Item(2).VerticalAlignment = [Microsoft.Office.Interop.Word.wdCellVerticalAlignment]::wdCellAlignVerticalCenter;
      $row.Cells.Item(2).Select()
      $selection.ClearParagraphAllFormatting()
      
      $pf = $selection.paragraphFormat
      $pf.LeftIndent = $word.CentimetersToPoints(0)
      $pf.RightIndent = $word.CentimetersToPoints(0)
      $pf.Alignment = [Microsoft.Office.Interop.Word.wdParagraphAlignment]::wdAlignParagraphRight
      $pf.SpaceBefore = 0
      $pf.SpaceBeforeAuto = $True
      $pf.SpaceAfter = 0
      $pf.SpaceAfterAuto = $True
      $pf.LineSpacingRule = [Microsoft.Office.Interop.Word.wdLineSpacing]::wdLineSpaceSingle
      $pf.CharacterUnitLeftIndent = 0
      $pf.CharacterUnitRightIndent = 0
      $pf.LineUnitBefore = 0
      $pf.LineUnitAfter = 0
    }
    
    $table.Columns.Item(2).PreferredWidthType = 2
    $table.Columns.Item(2).PreferredWidth = 1

    $table.Select()
    $selection.Previous().InsertParagraphBefore() # Выбираем предыдущую строку и вставляем перенос строки перед ней
    $selection.Next(5).InsertParagraphBefore() # Выбираем следующую строку и вставляем перенос строки перед ней
  }
  # Обычная таблица
  else {
  
    $table.AllowAutoFit = $True
    $table.AutoFitBehavior([Microsoft.Office.Interop.Word.wdAutoFitBehavior]::wdAutoFitWindow)

    $table.Select()
    $selection.ClearParagraphAllFormatting()
    $pf = $selection.ParagraphFormat
    $pf.LeftIndent = 0
    $pf.RightIndent = 0
    $pf.SpaceBefore = 0
    #$pf.Size = 12
    $pf.SpaceBeforeAuto = $False
    $pf.SpaceAfter = 0
    $pf.SpaceAfterAuto = $False

    # Добавление номеров столбцов
  
    $table.Rows[1].Select()
    $selection.InsertRowsBelow(1)
    for ($c = 1; $c -le $table.Columns.Count; $c++) {
      $table.Cell(2, $c).Range.Text = $c
    }
  
    # Выравнивание первых 2-х строк по середине
    for ($r = 1; $r -le 2; $r++) {
      $table.Rows.Item($r).Select()
      $pf = $selection.paragraphFormat
      $pf.Alignment = [Microsoft.Office.Interop.Word.wdParagraphAlignment]::wdAlignParagraphCenter
    }
  
  }
}
$finalRecommendations = $finalRecommendations + "`n - Проверьте разрыв таблиц между страницами"

Write-Host "Обновление стилей..."

$heading1namelocal = $doc.styles.Item([Microsoft.Office.Interop.Word.wdBuiltinStyle]::wdStyleHeading1).NameLocal
$heading2namelocal = $doc.styles.Item([Microsoft.Office.Interop.Word.wdBuiltinStyle]::wdStyleHeading2).NameLocal
$headers = $heading1namelocal, "UnnumberedHeading1", "UnnumberedHeadingOne"
$nchapters = 0
$nfigures = 0
$nreferences = 0
$nappendices = 0

foreach ($par in $doc.Paragraphs) {
  $isHeading = $True
  $namelocal = $par.Range.CharacterStyle.NameLocal
  if ($namelocal -eq "UnnumberedHeadingOne") {
    $par.Style = $UnnumberedHeading1
  }
  elseif ($namelocal -eq "AppendixHeadingOne") {
    $par.Style = $UnnumberedHeading1
    $nappendices = $nappendices + 1
  }
  elseif ($namelocal -eq "UnnumberedHeadingOneNoTOC") {
    $par.Style = $UnnumberedHeading1NoTOC
  }
  elseif ($namelocal -eq "UnnumberedHeadingTwo") {
    $par.Style = $UnnumberedHeading2
  }
  else {
    $namelocal = $par.Style.NameLocal
    # Make source core paragraphs smaller
    if ($namelocal -eq "Source Code") {
      $isHeading = $False
      $par.Range.Font.Size = 12
    }
    # No special style for first paragraph to avoid unwanted space
    # between first and second paragraphs
    elseif ($namelocal -eq "First Paragraph") {
      $isHeading = $False
      $par.Style = $BodyText
    }
    elseif ($namelocal -eq $heading1namelocal) {
      $nchapters = $nchapters + 1
    }
    elseif ($namelocal -eq $heading2namelocal) {
      $isHeading = $True
    }
    elseif ($namelocal -eq "Captioned Figure") {
      $isHeading = $False
      $nfigures = $nfigures + 1
    }
    elseif ($namelocal -eq "ReferenceItem") {
      $isHeading = $False
      $nreferences = $nreferences + 1
    }
    else{
      $isHeading = $False
    }    
    
    # Если нашли заголовок
    if ($isHeading){
      # Убираем пробелы перед заголовками
      $doc.Range($par.Range.Start, $par.Range.Start + 1).Select() | out-null
      if ($selection.Text -eq " "){
        $selection.Delete() | out-null
      }

      $doc.Range($par.Range.Start, $par.Range.Start).Select()
      # Добавляем разрыв страницы перед заголовками
      if ($namelocal -in $headers){
        $selection.InsertBreak([Microsoft.Office.Interop.Word.wdBreakType]::wdPageBreak)
      }
    }
  }
}

if ($counters) {
  Write-Host "Вставка количества глав, рисунков и таблиц..."
  $selection.HomeKey([Microsoft.Office.Interop.Word.wdUnits]::wdStory) | out-null
  $selection.Find.Execute("%NCHAPTERS%", $True, $True, $False, $False, $False, $True, `
      [Microsoft.Office.Interop.Word.wdFindWrap]::wdFindContinue, $False, $nchapters + "", `
      [Microsoft.Office.Interop.Word.wdReplace]::wdReplaceOne) | out-null
  $selection.HomeKey([Microsoft.Office.Interop.Word.wdUnits]::wdStory) | out-null
  $selection.Find.Execute("%NFIGURES%", $True, $True, $False, $False, $False, $True, `
      [Microsoft.Office.Interop.Word.wdFindWrap]::wdFindContinue, $False, $nfigures + "", `
      [Microsoft.Office.Interop.Word.wdReplace]::wdReplaceOne) | out-null
  $selection.HomeKey([Microsoft.Office.Interop.Word.wdUnits]::wdStory) | out-null
  $selection.Find.Execute("%NTABLES%", $True, $True, $False, $False, $False, $True, `
      [Microsoft.Office.Interop.Word.wdFindWrap]::wdFindContinue, $False, $ntables + "", `
      [Microsoft.Office.Interop.Word.wdReplace]::wdReplaceOne) | out-null
  $selection.HomeKey([Microsoft.Office.Interop.Word.wdUnits]::wdStory) | out-null
  $selection.Find.Execute("%NREFERENCES%", $True, $True, $False, $False, $False, $True, `
      [Microsoft.Office.Interop.Word.wdFindWrap]::wdFindContinue, $False, $nreferences + "", `
      [Microsoft.Office.Interop.Word.wdReplace]::wdReplaceOne) | out-null
  $selection.HomeKey([Microsoft.Office.Interop.Word.wdUnits]::wdStory) | out-null
  $selection.Find.Execute("%NAPPENDICES%", $True, $True, $False, $False, $False, $True, `
      [Microsoft.Office.Interop.Word.wdFindWrap]::wdFindContinue, $False, $nappendices + "", `
      [Microsoft.Office.Interop.Word.wdReplace]::wdReplaceOne) | out-null
}
       
Write-Host "Изменение размера математических формул..."

foreach ($math in $doc.OMaths) {
  # Size equations up a bit to match Paratype font size
  $math.Range.Font.Size = 14
}

Write-Host "Handling INCLUDEs..."
$selection.HomeKey([Microsoft.Office.Interop.Word.wdUnits]::wdStory) | out-null

while ($selection.Find.Execute("%INCLUDE(*)%^13", $True, $True, $True, $False, $False, $True, `
      [Microsoft.Office.Interop.Word.wdFindWrap]::wdFindContinue, $False, "", `
      [Microsoft.Office.Interop.Word.wdReplace]::wdReplaceNone)) {
  if ($selection.Text -match '%INCLUDE\((.*)\)%') {
    $filename = $matches[1]
    
    $start = $Selection.Range.Start
    $Selection.InsertFile([System.IO.Path]::GetFullPath($filename))
    
    if (!$?) {
      break
    }
      
    $end = $Selection.Range.End

    # Check if there is anything after the inserted documnt
    $selection.WholeStory()
    $totalend = $Selection.Range.End

    # If there is nothing after the inserted documnt, remove the extra CR which
    # mystically appears out of nowhere in that case
    if ($end -ge ($totalend - 1)) {
      $selection.Collapse([Microsoft.Office.Interop.Word.wdCollapseDirection]::wdCollapseEnd)  | out-null
      $selection.MoveLeft([Microsoft.Office.Interop.Word.wdUnits]::wdCharacter, 1, `
          [Microsoft.Office.Interop.Word.wdMovementType]::wdExtend)  | out-null
      $selection.Delete() | out-null
    }
  }
}

$selection.HomeKey([Microsoft.Office.Interop.Word.wdUnits]::wdStory) | out-null

Write-Host "Вставка оглавления..."
if ($selection.Find.Execute("%TOC%^13", $True, $True, $False, $False, $False, $True, `
      [Microsoft.Office.Interop.Word.wdFindWrap]::wdFindContinue, $False, "", `
      [Microsoft.Office.Interop.Word.wdReplace]::wdReplaceNone)) {
  $doc.TablesOfContents.Add($selection.Range, $False, 9, 9, $False, "", $True, $True, "", $True) | out-null
  
  # Manually add level 1,2,3 headers to ToC
  $toc = $doc.TablesOfContents.Item(1)
  $toc.UseHeadingStyles = $True
  #$toc.HeadingStyles.Add($UnnumberedHeading1, 1) | out-null
  #$toc.HeadingStyles.Add($UnnumberedHeading2, 2) | out-null
  $toc.HeadingStyles.Add($Heading1, 1) | out-null
  $toc.HeadingStyles.Add($Heading2, 2) | out-null
  $toc.HeadingStyles.Add($Heading3, 3) | out-null
  $toc.Update() | out-null

  # Изменим стиль оглавления
  $toc.Range.Font.Size = 14
  $toc.Range.Font.Name = "Times New Roman"
  $toc.Range.ParagraphFormat.LineSpacingRule = [Microsoft.Office.Interop.Word.wdLineSpacing]::wdLineSpace1pt5
  $toc.Range.ParagraphFormat.LeftIndent = 0

  $toc.Range.Font.SetAsTemplateDefault()

  # Добавим заголовок оглавления
  $tocTitle = "Оглавление"
  $toc.Range.InsertParagraphBefore()
  $toc.Range.Previous(4).Select()
  $selection.Text = $tocTitle + $selection.Text
  $selection.ParagraphFormat.Alignment = [Microsoft.Office.Interop.Word.wdParagraphAlignment]::wdAlignParagraphCenter
  $selection.Range.Bold = $True

  # Заменить разрыв страницы на разырыв "Следующая страница"
  $toc.Range.Next().Select()
  $selection.InsertBreak([Microsoft.Office.Interop.Word.wdBreakType]::wdSectionBreakNextPage)
  $doc.Range($selection.Range.Start, $selection.Range.End + 2).Select()
  $selection.Delete() | out-null
}

Write-Host "Нумерация страниц"

$Section = $Doc.Sections.Last
$Section.Footers.Item(1).LinkToPrevious = $False
$Section.Footers.Item(1).PageNumbers.Add(1) | out-null

Write-Host "Вставка количества страниц..."

# Seemingly is not needed but who knows
$doc.Repaginate()

# Inserting "section pages" field gets the number of pages wrong, and no way has
# been found to remedy that other than manual update in Word.
# So here is another way to get the number of pages in the section

if ($doc.Sections.Count -gt 1) {
  # two-section template?
  $npages = $doc.Sections.Item(2).Range.Information([Microsoft.Office.Interop.Word.wdInformation]::wdActiveEndPageNumber) - `
    $doc.Sections.Item(1).Range.Information([Microsoft.Office.Interop.Word.wdInformation]::wdActiveEndPageNumber)
}
else {
  $npages = $doc.Sections.Item(1).Range.Information([Microsoft.Office.Interop.Word.wdInformation]::wdNumberOfPagesInDocument)
}

$selection.HomeKey([Microsoft.Office.Interop.Word.wdUnits]::wdStory) | out-null
$selection.Find.Execute("%NPAGES%", $True, $True, $False, $False, $False, $True, `
    [Microsoft.Office.Interop.Word.wdFindWrap]::wdFindContinue, $False, $npages + "", `
    [Microsoft.Office.Interop.Word.wdReplace]::wdReplaceOne) | out-null

if ($embedfonts) {
  # Embed fonts (for users who do not have Paratype fonts installed).
  # This costs a few MB in file size
  $word.ActiveDocument.EmbedTrueTypeFonts = $True
  $word.ActiveDocument.DoNotEmbedSystemFonts = $True
  $word.ActiveDocument.SaveSubsetFonts = $True 
}

if (-not $is_docx_temporary) {
  Write-Host "Сохранение документа..."
  $doc.Save()
}

$doc.Close()
$word.Quit()

Write-Host "Удаление временных файлов..."
Remove-item -path $tempdocx
if ($is_docx_temporary) {
  Remove-item -path $docx
}

Write-Host `n$finalRecommendations