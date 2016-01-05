param([string]$ReplaceFileList=".\replaceFile.lst", [switch]$Force, [switch]$Debug, [switch]$Dryrun, [string]$TmplTitleSheet=".\表紙見本.xlsx")

# Some Const..

# Change SpecialChar as normal string to Valid one.
# TODO: This function move to shared library.
function Validate-SpecialChar ($str) {
	#TODO: handle str as ref for performance?
	$str = $str.replace('`n',"`n")
	$str = $str.replace('`r',"`r")
	$str = $str.replace('`t',"`t")
	$str = $str.replace('`a',"`a")
	#TODO: replace single & double qoute
	#$str = $str.replace('`''',"`'")
	#$str = $str.replace('`"',"`\"")
	$str = $str.replace('`0',"`0")
	$str = $str.replace('``',"``")
	return $str
}

# This is General-Purpose function to change "from" string to "to" string in "range".
function Replace-From-To ($range, $from, $to, [switch]$regexp, [bool]$dryrun) {
	$from = Validate-SpecialChar $from
	$to = Validate-SpecialChar $to
	$newOne = $range | where { $_.Text -eq $to }
	if ( -not $newOne )	{
		if ( $regexp ) {
			$cell = ($range | where { $_.Text -match $from })
		} else {
			$cell = ($range | where { $_.Text -eq $from })
		}
		if ( $cell ) {
			if ( $dryrun ) {
				Write-Host "DRYRUN: "$cell.Text " should change to $to"
			} else {
				$cell.Value() = "$to"
			}
		} else {
			echo "change fail $from to $to"
		}
	} else {
		echo "already changed to $to"
	} 
}

function Replace-TitleSheet-From-Tsv ($xl, $line, $TmplTitleSheet, [bool]$dryrun=$False) {
	# skip comment line starting #
	if ( $line -match "^#.*" ) {
		continue
	} 

	echo "$line"
	# filename,titleFrom,titleTo,dateFrom,dateTo,orgFrom,orgTo
	$line_array = $line.split("`t")
	#TODO: Use Destructuring.
	$filename = $line_array[0]
	$titleFrom = $line_array[1]
	$titleTo = $line_array[2]
	$versionFrom = $line_array[3]
	$versionTo = $line_array[4]
	$dateFrom = $line_array[5]
	$dateTo = $line_array[6]
	$orgFrom = $line_array[7]
	$orgTo = $line_array[8]

	# Change relative PATH to absolute PATH because Com run on CWD as User's HOME. 
	if ( $filename -notmatch "^[A-Za-z]:\\" ) {
		$filename = (Get-Location).Path + '\' + $filename
	}

	$wb = $xl.Workbooks.Open($filename)
	
	try {
		if ( $Force ) {
			# Force Replace of Title Sheet
			echo "Force replace 表紙シート"
			if ( $TmplTitleSheet -notmatch "^[A-Za-z]:\\" ) {
				$TmplTitleSheet = (Get-Location).Path + '\' + $TmplTitleSheet
			}
			$wb_tmpl = $xl.Workbooks.Open($TmplTitleSheet)
			$srcTitleSheet = ($wb_tmpl.WorkSheets | where { $_.name -eq "表紙" })
	        if ( -not $srcTitleSheet ) {
				throw "No Source Title Sheet"
	        } 
			$titleSheet = ($wb.Worksheets | where { $_.name -eq "表紙" })
			if ( $Dryrun ) {
				if ( -not $titleSheet ) {
					echo "表紙 Sheet not exist in $filename."
				}
				echo "DRYRUN: $TmplTitleSheet 's 表紙 sheet should copy to $filename. return here."
				return 
			} else {
				if ( $titleSheet ) {
					$titleSheet.Delete()
				} else {
					echo "表紙 Sheet not exist."
				}
				$srcTitleSheet.copy($wb.WorkSheets.Item(1))
			}
			$titleFrom = "TITLE_SOURCE"
			$versionFrom = "VERSION_SOURCE"
			$dateFrom = "DATE_SOURCE"
			$orgFrom = "ORGANIZATION_SOURCE"
		} 
		echo "filename is $filename"
		echo "titleFrom is $titleFrom"
		echo "titleTo is $titleTo"
		echo "versionFrom is $versionFrom"
		echo "versionTo is $versionTo"
		echo "dateFrom is $dateFrom"
		echo "dateTo is $dateTo"
		echo "orgFrom is $orgFrom"
		echo "orgTo is $orgTo"
		# Replace when *From value exists.
		$titleSheet = ($wb.Worksheets | where { $_.name -eq "表紙" })
		$titleSheetRange = $titleSheet.Range("A1:K12")
		echo "Replace Title"
		Replace-From-To $titleSheetRange $titleFrom $titleTo -dryrun $dryrun
		echo "Replace Version"
		Replace-From-To $titleSheetRange $versionFrom $versionTo -regexp -dryrun $dryrun
		echo "Replace Date"
		Replace-From-To $titleSheetRange $dateFrom $dateTo -regexp -dryrun $dryrun
		echo "Replace Organization"
		Replace-From-To $titleSheetRange $orgFrom $orgTo -dryrun $dryrun
	} catch [Exception] {
		echo $error[0].exception        
	} finally {
		$wb.Close($True)
		if ( $wb_tmpl ) {
			$wb_tmpl.Close($True)
	    }
	}
}

$xl = new-object -comobject excel.application

try {
	if ( $Debug ) {
		$xl.Visible = $True
		$xl.DisplayAlerts = $True
	} else {
		$xl.Visible = $False
		$xl.DisplayAlerts = $False
	}

    if ( $Dryrun ) {
		$dryrunBool = $True
	} else {
		$dryrunBool = $False
	}
	
	$repListFile = (Get-Content $ReplaceFileList) -as [string[]]
	
	foreach ( $line in $repListFile ) {
		Replace-TitleSheet-From-Tsv $xl $line $TmplTitleSheet -dryrun $DryrunBool
	}
} catch [Exception] {
	echo $error[0].exception        
} finally {
	$xl.Quit()
	[System.Runtime.Interopservices.Marshal]::ReleaseComObject($xl)
}

# vim: tw=0 tabstop=4:
