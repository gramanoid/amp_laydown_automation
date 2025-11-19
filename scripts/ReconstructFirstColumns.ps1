param(
    [Parameter(Mandatory=$true)][string]$PresentationPath,
    [int]$SlideIndex = 2,
    [int]$MaxColumns = 3
)

$path = [System.IO.Path]::GetFullPath($PresentationPath)
$pp = New-Object -ComObject PowerPoint.Application
$pp.Visible = 0
$presentation = $pp.Presentations.Open($path,0,0,0)
$slide = $presentation.Slides.Item($SlideIndex)
$table = $slide.Shapes.Item("MainDataTable").Table

$rows = @()
for($r=1;$r -le $table.Rows.Count;$r++){
    $rowTexts = @()
    for($c=1;$c -le $MaxColumns;$c++){
        try{
            $rowTexts += $table.Cell($r,$c).Shape.TextFrame.TextRange.Text
        }catch{
            $rowTexts += ""
        }
    }
    $rows += ,[pscustomobject]@{Row=$r;Texts=$rowTexts}
}

for($i=1;$i -le $MaxColumns;$i++){
    $table.Columns.Delete(1)
}

for($i=1;$i -le $MaxColumns;$i++){
    $null = $table.Columns.Add(1,$i)
}

foreach($row in $rows){
    $rowIndex = $row.Row
    $texts = $row.Texts
    for($c=1;$c -le $MaxColumns;$c++){
        $table.Cell($rowIndex,$c).Shape.TextFrame.TextRange.Text = $texts[$c-1]
    }
}

$presentation.Save()
$presentation.Close()
$pp.Quit()
