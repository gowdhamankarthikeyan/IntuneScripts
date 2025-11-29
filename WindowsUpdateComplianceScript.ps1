$Script:RegPath = "HKLM:\Software\WindowsUpdateCompliance\"
Add-Type -AssemblyName System.Web
$StatusTime = Get-Date

function Get-Windows11ReleaseTableContent {
    $url = "https://learn.microsoft.com/en-us/windows/release-health/windows11-release-information"
    try {
        $response = Invoke-WebRequest -Uri $url -UseBasicParsing
        $html = [System.Web.HttpUtility]::HtmlDecode($response.Content)
        $rows = @()
        $index = 1
		# Basic fallback: Split on <table> tags (less reliable for nested content)
		$tableMatches = [regex]::Matches($html, '<table[^>]*>.*?</table>', [System.Text.RegularExpressions.RegexOptions]::Singleline)
		foreach ($match in $tableMatches) {
			$tableHtml = $match.Value
			# Extract title (heuristic)
			$title = if ($tableHtml -match '<caption[^>]*>(.*?)</caption>') { $matches[1].Trim() } else { "Table $index" }

			# Headers (first tr with th/td)
			$headers = @()
			if ($tableHtml -match '<tr[^>]*>.*?(<th[^>]*>.*?</th>|<td[^>]*>.*?</td>).*?</tr>') {
				$rowContent = $matches[0]
				$thMatches = [regex]::Matches($rowContent, '<th[^>]*>(.*?)</th>|<td[^>]*>(.*?)</td>')
				foreach ($thMatch in $thMatches) { $headers += ($thMatch.Groups[1].Value + $thMatch.Groups[2].Value).Trim() }
			}
			If ($headers -match "Update type"){
				# Rows (subsequent tr)
				$rowMatches = [regex]::Matches($tableHtml, '(?<=<tr[^>]*>).*?(?=</tr>)', [System.Text.RegularExpressions.RegexOptions]::Singleline)
				for ($i = 1; $i -lt $rowMatches.Count; $i++) {  # Skip first (header)
					$rowContent = $rowMatches[$i].Value
					$cellMatches = [regex]::Matches($rowContent, '<td[^>]*>(.*?)</td>|<th[^>]*>(.*?)</th>')
					$rowData = [ordered]@{}
					for ($j = 0; $j -lt [Math]::Min($headers.Count, $cellMatches.Count); $j++) {
						$cellText = ($cellMatches[$j].Groups[1].Value + $cellMatches[$j].Groups[2].Value).Trim()
						# Basic link extract
						if ($cellText -match 'href=["'']([^"'']+)["''][^>]*>([^<]+)</a>') {
							#$cellText = "$($matches[2].Trim()) ($($matches[1].Trim()))"
							$rowData['KB article'] = $matches[2].Trim()
							$rowData['Support Link'] = $matches[1].Trim()
						} Else {
							If ($cellText -match "<span>.*</span>") { $cellText = $cellText -replace '<span>.*</span>','-'}
							$rowData[$headers[$j]] = $cellText
						}
						
					}
					if ($rowData.Count -gt 0) { $rows += $rowData }
				}
			}
			$index++
		}
        return $rows
    }
    catch {
        Write-Error "Fetch/parse error: $($_.Exception.Message)"
        return @()
    }
}

function Get-Windows10ReleaseTableContent {
    $url = "https://learn.microsoft.com/en-us/windows/release-health/release-information"
    try {
        $response = Invoke-WebRequest -Uri $url -UseBasicParsing
        $html = [System.Web.HttpUtility]::HtmlDecode($response.Content)
        $rows = @()
        $index = 1
		# Basic fallback: Split on <table> tags (less reliable for nested content)
		$tableMatches = [regex]::Matches($html, '<table[^>]*>.*?</table>', [System.Text.RegularExpressions.RegexOptions]::Singleline)
		foreach ($match in $tableMatches) {
			$tableHtml = $match.Value
			# Extract title (heuristic)
			$title = if ($tableHtml -match '<caption[^>]*>(.*?)</caption>') { $matches[1].Trim() } else { "Table $index" }

			# Headers (first tr with th/td)
			$headers = @()
			if ($tableHtml -match '<tr[^>]*>.*?(<th[^>]*>.*?</th>|<td[^>]*>.*?</td>).*?</tr>') {
				$rowContent = $matches[0]
				$thMatches = [regex]::Matches($rowContent, '<th[^>]*>(.*?)</th>|<td[^>]*>(.*?)</td>')
				foreach ($thMatch in $thMatches) { $headers += ($thMatch.Groups[1].Value + $thMatch.Groups[2].Value).Trim() }
			}
			If ($headers -match "Update type"){
				# Rows (subsequent tr)
				$rowMatches = [regex]::Matches($tableHtml, '(?<=<tr[^>]*>).*?(?=</tr>)', [System.Text.RegularExpressions.RegexOptions]::Singleline)
				for ($i = 1; $i -lt $rowMatches.Count; $i++) {  # Skip first (header)
					$rowContent = $rowMatches[$i].Value
					$cellMatches = [regex]::Matches($rowContent, '<td[^>]*>(.*?)</td>|<th[^>]*>(.*?)</th>')
					$rowData = [ordered]@{}
					for ($j = 0; $j -lt [Math]::Min($headers.Count, $cellMatches.Count); $j++) {
						$cellText = ($cellMatches[$j].Groups[1].Value + $cellMatches[$j].Groups[2].Value).Trim()
						# Basic link extract
						if ($cellText -match 'href=["'']([^"'']+)["''][^>]*>([^<]+)</a>') {
							#$cellText = "$($matches[2].Trim()) ($($matches[1].Trim()))"
							$rowData['KB article'] = $matches[2].Trim()
							$rowData['Support Link'] = $matches[1].Trim()
						} Else {
							$rowData[$headers[$j]] = $cellText
						}
						
					}
					if ($rowData.Count -gt 0) { $rows += $rowData }
				}
			}
			$index++
		}
        return $rows
    }
    catch {
        Write-Error "Fetch/parse error: $($_.Exception.Message)"
        return @()
    }
}

#Create Loging Registry Path
If (!(Test-Path $RegPath)){
	New-Item -Path $RegPath -Force -ErrorAction SilentlyContinue | Out-Null
}

#Get OS Build Number
$ComputerInfo = Get-ComputerInfo
$OSDisplayVersion = $ComputerInfo.OSDisplayVersion
$OSName = $ComputerInfo.OSName
$WinCV = (Get-ItemProperty -Path "HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion")
$OSBuildNumber = $WinCV.CurrentBuild + "." + $WinCV.UBR
$CurrentUpdate = [ordered]@{}
$DaysSinceCurrentUpdateReleaseDate = -1

If ($OSName -match "Windows 11") {
	#Windows 11
	$Updates = Get-Windows11ReleaseTableContent
} Elseif ($OSName -match "Windows 10") {
	#Windows 10
	$Updates = Get-Windows10ReleaseTableContent
} Else {
	#Neither Windows 10 Nor Windows 11
	$Status = "NeitherWindows10NotWindows11"
}

If ($Status -ne "NeitherWindows10NotWindows11"){
	$CurrentUpdate = $Updates | where {$_.build -eq $OSBuildNumber}
	If ($CurrentUpdate){
		$Status = "UpdateFoundInCatalog"
		$Properties = [PSCustomObject]$CurrentUpdate| Get-Member -MemberType NoteProperty
		foreach ($Property in $Properties){
			$PropertyName = $($Property.Name)
			$PropertyValue = $CurrentUpdate.$PropertyName
			New-ItemProperty -Path $RegPath -Name $PropertyName -Value $PropertyValue -PropertyType String -Force -ErrorAction SilentlyContinue | Out-Null
		}
		$AvailabilityDate = Get-Date ($CurrentUpdate.'Availability date')
		$CurrentDate = Get-Date
		#Rounding down to lowest integer
		$DaysSinceCurrentUpdateReleaseDate = [Math]::floor(($CurrentDate - $AvailabilityDate).TotalDays)
	} Else {
			$Status = "UpdateNotFoundInCatalog"
	}
} Else {
	$Status = "InsiderBuild"
}

#Formating Results
$CurrentUpdate['OSName'] = $OSName
$CurrentUpdate['OSBuildNumber'] = $OSBuildNumber
$CurrentUpdate['OSDisplayVersion'] = $OSDisplayVersion
$CurrentUpdate['DaysSinceCurrentUpdateReleaseDate'] = $DaysSinceCurrentUpdateReleaseDate
$CurrentUpdate['Status'] = $Status

New-ItemProperty -Path $RegPath -Name "OSName" -Value $OSName -PropertyType String -Force -ErrorAction SilentlyContinue | Out-Null
New-ItemProperty -Path $RegPath -Name "OSBuildNumber" -Value $OSBuildNumber -PropertyType String -Force -ErrorAction SilentlyContinue | Out-Null
New-ItemProperty -Path $RegPath -Name "OSDisplayVersion" -Value $OSDisplayVersion -PropertyType String -Force -ErrorAction SilentlyContinue | Out-Null
New-ItemProperty -Path $RegPath -Name "DaysSinceCurrentUpdateReleaseDate" -Value $DaysSinceCurrentUpdateReleaseDate -PropertyType String -Force -ErrorAction SilentlyContinue | Out-Null
New-ItemProperty -Path $RegPath -Name "Status" -Value $Status -PropertyType String -Force -ErrorAction SilentlyContinue | Out-Null
New-ItemProperty -Path $RegPath -Name "StatusTime" -Value $StatusTime -PropertyType String -Force -ErrorAction SilentlyContinue | Out-Null

return $CurrentUpdate | ConvertTo-Json -Compress
