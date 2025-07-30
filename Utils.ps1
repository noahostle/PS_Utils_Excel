function Get-Package-ColumnMap($ws) {
	$columns = @{}
	for ($col=1; $col -le $ws.Dimension.End.Column; $col++) {
		$header = $ws.Cells[1, $col].Text
		if ($header) {$columns[$header] = $col}
	}
	return $columns
}


function Check-Column-Exists($ws, $name, $default="") {
    $nCol=0
	for ($col = 1; $col -le $ws.Dimension.Columns; $col++) {
        $value = $ws.Cells[1, $col].Text
        if ($value -eq $name) {
           	$nCol=$col
		    break
       }
    } 

	if ($nCol -eq 0) {
		$nCol = $ws.Dimension.Columns + 1
		$head = $ws.Cells[1, $nCol]
		$head.Style.Fill.PatternType = 'Solid'
		$head.Style.Fill.BackgroundColor.SetColor([System.Drawing.Color]::LightBlue)
		$head.Style.Font.Bold = $true
		$head.Value = $name
		for ($x=2; $x -le $ws.Dimension.Rows; $x++) {
			$ws.Cells[$x, $nCol].Value = $default
		}
	}

	return $nCol
}



function Get-SheetName($dir) {
	# Requires ImportExcel module
	if (-not (Get-Module -ListAvailable ImportExcel)) { Install-Module ImportExcel -Scope CurrentUser -Force }
	Import-Module ImportExcel -Force

	# Get all .xlsx files in current directory
	$excelFiles = Get-ChildItem -Path $dir -Filter *.xlsx
	
	Write-Host ""

	if ($excelFiles.Count -eq 0) {
		Write-Error "No Excel files found in the current directory." 
		exit
	}
	elseif ($excelFiles.Count -eq 1) {
		$filename = $excelFiles[0].FullName
		Write-Host "Using: $filename" -ForegroundColor Green
	}
	else {
		# Prompt user to pick a file
		Write-Host "Please select a .xlsx file:" -ForegroundColor Yellow
		for ($i = 0; $i -lt $excelFiles.Count; $i++) {
			Write-Host "  [$i] $($excelFiles[$i].Name)"
		}
		do {
			Write-Host "`n>> " -NoNewLine
			$selection = $Host.UI.ReadLine()
		} while (-not ($selection -match '^\d+$') -or $selection -lt 0 -or $selection -ge $excelFiles.Count)

		$filename = $excelFiles[$selection].FullName
	}

	Write-Host ""
	
	# Get sheet names from the chosen Excel file
	$sheets = (Get-ExcelSheetInfo -Path $filename).Name
	if ($sheets.Count -eq 0) {
		Write-Error "No sheets found in the Excel file." 
		exit
	}
	
	if ($sheets -is [string]) {
		return @($filename, $sheets)
	}

	Write-Host "Available sheets:" -ForegroundColor Yellow
	for ($i = 0; $i -lt $sheets.Count; $i++) {
		Write-Host "  [$i] $($sheets[$i])"
	}

	do {
		Write-Host "`n>> " -NoNewLine
		$sheetSelection = $Host.UI.ReadLine()
	} while (-not ($sheetSelection -match '^\d+$') -or $sheetSelection -lt 0 -or $sheetSelection -ge $sheets.Count)

	$sheetName = $sheets[$sheetSelection]

	Write-Host ""

	return @($filename, $sheetName)
}

function Get-SheetData($dir) {
	$result = Get-SheetName $dir
	# Import data from selected sheet into $data
	$data = Import-Excel -Path $result[0] -WorksheetName $result[1]
	return $data
}

function Get-SheetPKG($dir) {
	$result = Get-SheetName $dir
	$pkg = Open-ExcelPackage -Path $result[0]
	$ws = $pkg.Workbook.Worksheets[$result[1]]
	return @($pkg, $ws, $result)
}
