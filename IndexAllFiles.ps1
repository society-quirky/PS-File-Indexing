# Produce CSV index files:
#	1. For all documents within root folder
# 	2. For all files with "keywords" (as phrase) in title
#	3. For all files with "TRANSCRIPT" in title (with Date Time yyyymmddhhmm in filename)
#	Columns: Document name | Date (DD Mon YYYY) | Reference | Source file path (trimmed)
#	Save outputs to specified folder

param ()

# Configuration
# Replace words in [*] below with relevant file path
# $RootPath for folder from which all files are to be listed
# $TrimPrefix can be defined as the path to the root folder 

$RootPath 	= '[FILE PATH]'
$TrimPrefix 	= '[TRIMMED PREFIX TO PRINT]'
$OutDir		= '[WHERE TO SAVE OUTPUT]'

# Setup and Validate Paths

$culture = [System.Globalization.CultureInfo]::GetCultureInfo('en-UK')

Write-host "Validating paths..." -ForegroundColor Cyan
if (-not (Test-Path -LiteralPath $RootPath)) {
	Write-Error "Root path not found or inaccessible: $RootPath"
	exit 1
}

if (-not (Test-Path -LiteralPath $OutDir)) {
	Write-Host "Creating output directory: $OutDir"
	New-Item -ItemType Directory -Path $OutDir -Force | Out-Null
}

# Read file name for date references

function Normalize-MonthText {
	param([string]$mon)
	if ([string]::IsNullOrWhiteSpace($mon)) { return $null }

	$m = $mon
	if ($m.Length -gt 5) { $m = $m.Substring(0,5) }
	$m = $m.ToLowerInvariant()

	switch ($m) {
	'jan' { return 1 }
	'feb' { return 2 }
	'mar' { return 3 }
	'apr' { return 4 }
	'may' { return 5 }
	'jun' { return 6 }
	'jul' { return 7 }
	'aug' { return 8 }
	'sep' { return 9 }
	'sept' { return 9 }
	'oct' { return 10 }
	'nov' { return 11 }
	'dec' { return 12 }
	default {
		if ($mon -match '^\d{1,2}$') {
		    $n = [int]$mon
		    if ($n -ge 1 -and $n -le 12) { return $n }
		}
		return $null
	    }
	}
}

function Format-DateDDMonYYYY {
	param([int]$year, [int]$month, [int]$day)
	try {
	    $dt = Get-Date -Year $year -Month $month -Day $day -Hour 0 -Minute 0 -Second 0
	    return $dt.ToString('dd MMM yyyy', $culture)
	} catch {return ''}
}

# Extract a 6-12 digit reference number from the relative path, else blank
function Get-ReferenceNo {
	param([string]$RelativePath)
	if ([string]::IsNullOrWhiteSpace($RelativePath)) { return '' }
	$m = [System.Text.RegularExpressions.Regex]::Match($RelativePath, '(?<!\d)(\d{6,12})(?!\d)')
	if ($m.Success) { return $m.Groups[1].Value }
	return ''
}
# Extract date from filename , return 'DD MMM YYYY' or empty string.
function Get-DateFromName {
	param([string]$Name)
	if ([string]::IsNullOrWhiteSpace($Name)) { return ''}

	#Take the LAST match found among the patterns
	$patterns = @(
	 # 1) D(D)? Mon Y(YY|YYY)Y with separators space/_/-/., e.g. '1 Sep 23', '15 Sept 2025' 
	'(?<!\d)(?<day>\d{1,2})[\s_\-\.]+(?<mon>Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Sept|Oct|Nov|Dec)[\s_\-\.]+(?<year>\d{2,4})(?!\d)',
	# 2) YYYYMMDD
	'(?<!\d)(?<year>19\d{2}|20\d{2})(?<mon>\d{2})(?<day>\d{2})(?!\d)',
	# 3) DDMMYYYY
	'(?<!\d)(?<day>\d{2})(?<mon>\d{2})(?<year>19\d{2}|20\d{2})(?!\d)',
	# 4) YYYY[-_. ]MM[-_. ]DD
	'(?<!\d)(?<year>19\d{2}|20\d{2})[\-_\.\s](?<mon>\d{2})[\-_\.\s](?<day>\d{2})(?!\d)',
	# 5) DD[-_. ]MM[-_. ]YYYY
	'(?<!\d)(?<day>\d{2})[\-_\.\s](?<mon>\d{2}[\-_\.\s])(?<year>19\d{2}|20\d{2})(?!\d)'
)
	foreach ($rx in $patterns) {
		$matches = [System.Text.RegularExpressions.Regex]::Matches($Name, $rx,
[System.Text.RegularExpressions.RegexOptions]::IgnoreCase)
	if ($matches.Count -gt 0) {
	    $m = $matches[$matches.Count - 1] # last match

	    $dayStr 	= $m.Groups['day'].Value
	    $monStr 	= $m.Groups['mon'].Value
	    $yearStr 	= $m.Groups['year'].Value
	# Parse day
	$day = 0
	[void][int]::TryParse($dayStr, [ref]$day)

	# Parse year (expand 2-digit)
	$year = 0
	if ([int]::TryParse($yearStr, [ref]$year)) {
	    if ($year -lt 100) {
		if ($year -ge 50) { $year = 1900 + $year } else { $year = 2000 + $year }
	    }
	}

	# Month: text or numeric
	$month = $null
	if ($monStr -and ($monStr -match '[A-Za-z]')) {
	    $month = Normalize-MonthText $monStr
	} else {
	    $month = Normalize-MonthText $monStr
	}

	if ($day -ge 1-and $day -le 31 -and $year -ge 1900 -and $year -le 2099 -and $month -ge 1 -and $month -le 12) { return (Format-DateDDMonYYYY -year $year -month $month -day $day)
	}
    }
}
return ''
}

# Extract transcript datetime from filename as yyyymmmddhhmm (with or without separator).
# Return 'yyyy-MM-dd HH:mm' or empty if not found/invalid.
function Get-DateTimeFromFileName {
	param([string]$Name)
	if ([string]::IsNullOrWhiteSpace($Name)) { return ''}
	# last yyyymmdd[blc]hhmmm; blc can be space, '_', '-', or none
	$rx = '(?<!\d)(?<ymd>\d{8})[\s_\-]?(?<hm>\d{4})(?!\d)'
	$matches = [System.Text.RegularExpressions.Regex]::Matches($Name, $rx)
	if ($matches.Count -eq 0) { return '' }

	$m = $matches[$matches.Count - 1] #last occurrence
	$ymd = $m.Groups['ymd'].Value
	$hm  = $m.Groups['hm'].Value

	$year 	= [int]$ymd.Substring(0,4)
	$month  = [int]$ymd.Substring(4,2)
	$day	= [int]$ymd.Substring(6,2)
	$hour	= [int]$hm.Substring(0,2)
	$min	= [int]$hm.Substring(2,2)

	#VALIDATE
	if ($year -lt 1900 -or $year -gt 2099) { return '' }
	if ($month -lt 1 -or $month -gt 12) { return '' }
	if ($day -lt 1 -or $day -gt 31) { return '' }

	try {
		$dt = Get-Date -Year $year -Month $month -Day $day -Hour $hour -Minute $min -Second 0
		return $dt.ToString('yyyy-MM-dd HH:mm', $culture)
	} catch {
		return ''
	}
}

#Enumerate files
Write-host "Enumerating files under: $RootPath" -ForegroundColor Cyan
$files = Get-ChildItem -LiteralPath $RootPath -File -Recurse -ErrorAction SilentlyContinue

#Count files
$fileCount = ($files | Measure-Object).Count
Write-Host "Total files found: $fileCount" -ForegroundColor Green

#Build row (base)
$rows = foreach ($f in $files) {
	$relative = $f.FullName
	if ($relative.StartsWith($TrimPrefix, [System.StringComparison]::OrdinalIgnoreCase)) {
	    $relative = $relative.Substring($TrimPrefix.Length)
	}

	$docName = $f.Name
	$dateOut = Get-DateFromName -Name $docName
	$dateTime = Get-DateTimeFromFileName
	$refno 	 = Get-ReferenceNo -RelativePath $relative

	[pscustomobject]@{
		'Document name' 	= $docName
		'Date (DD Mon YYYY)'	= $dateOut
		'Date Time'		= $dateTime
		'Reference'		= $refno
		'Source file path'	= $relative
	}
}

#Filters
$allIndex 	= $rows
#Search exact phrases in file name as an exact phrase, use separator "[_\-]?" for space, e.g. '(?i)monthly[_\-]?statement'
#Search for any words in file name, use separator "|", e.g. '(?i)statement|report'
$keywordsIndex 	= $rows | Where-Object { $_.'Document name' -match '(?i)keyword1[ _\-]?keyword2' }
#Search for any single keword in file name, replace "transcript" in following
$transcriptIndex = $rows | Where-Object { $_.'Document name' -match '(?i)transcript' }

# Write outputs
$stamp = (Get-Date).ToString('yyyyMMdd-HHmmss')

$fileCountPath 	= Join-Path $OutDir ("FileCount_{0}.txt" -f $stamp)
$allCsvPath 	= Join-Path $OutDir ("Index_ALLFiles{0}.csv" -f $stamp)
$keywordsCsvPath = Join-Path $OutDir ("Index_Keywords_{0}.csv" -f $stamp)
$transcriptCsvPath = Join-Path $OutDir ("Index_Transcripts_{0}.csv" -f $stamp)

"Total files: $fileCount" | Out-File -FilePath $fileCountPath -Encoding UTF8
$allIndex 		| Sort-Object 'Document name' | Export-Csv -Path $allCsvPath -NoTypeInformation -Encoding UTF8
$keywordsIndex 		| Sort-Object 'Document name' | Export-Csv -Path $keywordsCsvPath -NoTypeInformation -Encoding UTF8
$transcriptIndex  	| Sort-Object 'Document name' | Export-Csv -Path $transcriptCsvPath -NoTypeInformation -Encoding UTF8 

Write-Host ""
Write-Host "Output files created:" -ForegroundColor Cyan
Write-Host " - $fileCountPath"
Write-Host " - $allCsvPath"
Write-Host " - $keywordsCsvPath"
Write-Host " - $transcriptCsvPath"
Write-Host ""
Write-Host "Done." -ForegroundColor Green