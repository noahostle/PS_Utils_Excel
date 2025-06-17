function Get-Sheet-From-PWD($dir) {
	# Requires ImportExcel module
	if (-not (Get-Module -ListAvailable ImportExcel)) { Install-Module ImportExcel -Scope CurrentUser -Force }
	Import-Module ImportExcel -Force

	# Get all .xlsx files in current directory
	$excelFiles = Get-ChildItem -Path $dir -Filter *.xlsx
	
	Write-Host ""

	if ($excelFiles.Count -eq 0) {
		Write-Error "No Excel files found in the current directory." -ForegroundColor Red
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
		Write-Error "No sheets found in the Excel file." -ForegroundColor Red
		exit
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

	# Import data from selected sheet into $data
	$data = Import-Excel -Path $filename -WorksheetName $sheetName
	

	Write-Host ""

	return $data
}
