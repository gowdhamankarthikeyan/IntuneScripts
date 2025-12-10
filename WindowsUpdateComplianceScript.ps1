Add-Type -AssemblyName System.Web
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

Function Get-AppliesTo {
	param (
		$Uri
	)
	try {
		$AppliesTo = $Null
		$response = Invoke-WebRequest -Uri $Uri -UseBasicParsing
		$html = [System.Web.HttpUtility]::HtmlDecode($response.Content)
		$AppliesToMatches = [regex]::Matches($html, '<span class="appliesToItem"[^>]*>.*?</span>', [System.Text.RegularExpressions.RegexOptions]::Singleline)
		foreach ($match in $AppliesToMatches) {
			$AppliesToMatchesHTML = $match.Value
			$AppliesTo += if ($AppliesToMatchesHTML -match '<span[^>]*>(.*?)</span>') { $matches[1].Trim() }
		}
		return $AppliesTo
	}
	catch {
        Write-Error "Fetch/parse error: $($_.Exception.Message)"
        return @()
    }
}

#Create Logging Registry Path
$Script:RegPath = "HKLM:\Software\WindowsUpdateCompliance\"
If (Test-Path $RegPath){
	Remove-Item -Path $RegPath -Force -ErrorAction SilentlyContinue | Out-Null
	New-Item -Path $RegPath -Force -ErrorAction SilentlyContinue | Out-Null
} Else {
	New-Item -Path $RegPath -Force -ErrorAction SilentlyContinue | Out-Null
}

#Get OS Build Number
$ComputerInfo = Get-ComputerInfo
$OSDisplayVersion = $ComputerInfo.OSDisplayVersion
$OSName = $ComputerInfo.OSName
$WinCV = (Get-ItemProperty -Path "HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion")
$OSBuildNumber = $WinCV.CurrentBuild + "." + $WinCV.UBR

#Initialize Variables
$StatusTime = Get-Date
$CurrentUpdate = [ordered]@{}
$DaysSinceCurrentUpdateReleaseDate = $DaysSinceLatestUpdateReleaseDate = -1
$CurrentPatchLevel = -1
$RequiredPatchLevel = 0
$QualityUpdateGracePeriod = $FeatureUpdateGracePeriod = $QualityUpdateDeadline = $FeatureUpdateDeadline = 0
$TotalGracePeriod = 0
$isLatestForThisDevice = $false
$12MonthsAgo = (Get-Date).AddMonths(-12)

#Get Windows Updates Policy Settings
# For Quality Update Grace Period
$QualityUpdateGracePeriod = (Get-ItemProperty -Path "HKLM:\SOFTWARE\Microsoft\PolicyManager\current\device\Update" -ErrorAction SilentlyContinue).ConfigureDeadlineGracePeriod
# For Feature Update Grace Period (if separately configured)
$FeatureUpdateGracePeriod = (Get-ItemProperty -Path "HKLM:\SOFTWARE\Microsoft\PolicyManager\current\device\Update" -ErrorAction SilentlyContinue).ConfigureDeadlineGracePeriodForFeatureUpdates
# For Quality Update Deadline (days from publication)
$QualityUpdateDeadline = (Get-ItemProperty -Path "HKLM:\SOFTWARE\Microsoft\PolicyManager\current\device\Update" -ErrorAction SilentlyContinue).ConfigureDeadlineForQualityUpdates
# For Feature Update Deadline (days from publication)
$FeatureUpdateDeadline = (Get-ItemProperty -Path "HKLM:\SOFTWARE\Microsoft\PolicyManager\current\device\Update" -ErrorAction SilentlyContinue).ConfigureDeadlineForFeatureUpdates


If ($OSName -match "Windows 11") {
	#Windows 11 Filter Updates applicable for Device's Build; Sort By Build Number
	$UpdatesForCurrentBuild = Get-Windows11ReleaseTableContent | where-object {($_.'Build' -match $($WinCV.CurrentBuild)) -and ($_.Type -notmatch ".")} | Sort-Object -Property @{e={[System.Version]($_.'Build')}} -Descending
	#Last 12 Months;
	$Updates = $UpdatesForCurrentBuild |where-object {$_.'Update type' -match 'B'} | where-object {[DateTime]($_.'Availability Date') -gt $12MonthsAgo}
	$ApplicableUpdates = @()
	Foreach ($Update in $Updates){
		If ($Update.Type -match "."){
			#Skip
		} Else {
			If ($Update.'Update type' -match 'OOB'){
				$AppliesTo = Get-AppliesTo -Uri $($Update.'Support Link')
				If ($AppliesTo -match "Windows 11"){ 
					$ApplicableUpdates += $Update
				}
			} Else { $ApplicableUpdates += $Update }
		}
	}
} Elseif ($OSName -match "Windows 10") {
	#Windows 10 Filter Updates applicable for Device's Build; Sort By Build Number
	$UpdatesForCurrentBuild = Get-Windows10ReleaseTableContent | where-object {($_.'Build' -match $($WinCV.CurrentBuild)) -and ($_.Type -notmatch ".")} | Sort-Object -Property @{e={[System.Version]($_.'Build')}} -Descending
	#Last 12 Months;
	$Updates = $UpdatesForCurrentBuild |where-object {$_.'Update type' -match 'B'} | where-object {[DateTime]($_.'Availability Date') -gt $12MonthsAgo}
	$ApplicableUpdates = @()
	Foreach ($Update in $Updates){
		If ($Update.Type -match "."){
			#Skip
		} Else {
			If ($Update.'Update type' -match 'OOB'){
				$AppliesTo = Get-AppliesTo -Uri $($Update.'Support Link')
				If ($AppliesTo -match "Windows 10"){ 
					$ApplicableUpdates += $Update
				}
			} Else { $ApplicableUpdates += $Update }
		}
	}
} Else {
	#Neither Windows 10 Nor Windows 11
	$Status = "NeitherWindows10NotWindows11"
}


If ($Status -ne "NeitherWindows10NotWindows11"){
	$CurrentUpdate = $UpdatesForCurrentBuild | where {$_.build -eq $OSBuildNumber}
	If ($CurrentUpdate){
		$CurrentUpdateAvailabilityDate = Get-Date ($CurrentUpdate.'Availability date')
		# Find Current Patch Level
		# Step 1 - Compare with $ApplicableUpdates
		$CurrentPatchLevel = $ApplicableUpdates.IndexOf($CurrentUpdate)
		$LatestApplicableUpdateAvailabilityDate = Get-Date ($ApplicableUpdates.'Availability date')[0]
		#Step 2 - If not, compare with $UpdatesForCurrentBuild (Most likely, the device is on D update or older than 12 months ago)
		If ($CurrentPatchLevel -eq "-1"){
			$CurrentPatchLevel = $UpdatesForCurrentBuild.IndexOf($CurrentUpdate)
			$LatestApplicableUpdateAvailabilityDate = Get-Date ($UpdatesForCurrentBuild.'Availability date')[0]
		}
		# If found, proceed
		If ($CurrentPatchLevel -ne "-1"){
			$Status = "UpdateFoundInCatalog"
			$CurrentDate = Get-Date
			#Rounding down to lowest integer
			$DaysSinceCurrentUpdateReleaseDate = [Math]::floor(($CurrentDate - $CurrentUpdateAvailabilityDate).TotalDays)
			$DaysSinceLatestUpdateReleaseDate = [Math]::floor(($CurrentDate - $LatestApplicableUpdateAvailabilityDate).TotalDays)
			#Find RequiredPatchLevel for the device incorporating the deferral and deadline values (If any)
			$TotalGracePeriod = $QualityUpdateGracePeriod + $QualityUpdateDeadline
			If ($DaysSinceLatestUpdateReleaseDate -lt $TotalGracePeriod) {
				# Device still has time to get latest applicable patch. So marking required as N-1
				$RequiredPatchLevel = 1
			} Else {
				# Time elapsed to get latest applicable patch. So marking required as N.
				$RequiredPatchLevel = 0
			}

			If ($CurrentPatchLevel -le $RequiredPatchLevel) { $isLatestForThisDevice = $true }
		}
	}
	Else {
		$Status = "InsiderBuild"
	}
}

# Format Result

If ($CurrentUpdate) {
	$Result = $CurrentUpdate
} Else {
	$Result = @{}
}

$Result['OSName'] = $OSName
$Result['OSBuildNumber'] = $OSBuildNumber
$Result['OSDisplayVersion'] = $OSDisplayVersion
$Result['DaysSinceCurrentUpdateReleaseDate'] = $DaysSinceCurrentUpdateReleaseDate
$Result['DaysSinceLatestUpdateReleaseDate'] = $DaysSinceLatestUpdateReleaseDate
$Result['Status'] = $Status
$Result['StatusTime'] = $StatusTime
$Result['CurrentPatchLevel'] = $CurrentPatchLevel
$Result['RequiredPatchLevel'] = $RequiredPatchLevel
$Result['isLatestForThisDevice'] = $isLatestForThisDevice
$Properties = [PSCustomObject] $Result| Get-Member -MemberType NoteProperty
foreach ($Property in $Properties){
	$PropertyName = $($Property.Name)
	$PropertyValue = $CurrentUpdate.$PropertyName
	New-ItemProperty -Path $RegPath -Name $PropertyName -Value $PropertyValue -PropertyType String -Force -ErrorAction SilentlyContinue | Out-Null
}

# Return Result
return $Result | ConvertTo-Json -Compress
