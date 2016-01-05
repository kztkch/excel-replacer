param([string]$ReplaceHFFileList=".\replaceHeaderFooterFile.lst", [switch]$Debug, [switch]$Dryrun)

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

function Replace-HeaderFooter ($xl, $configLine, $filenameLine, [bool]$dryrun=$False) {
	$configLine_array = $configLine.split("`t")
	#TODO: Use Destructuring.
	$hl = $configLine_array[0]
	$hc = $configLine_array[1]
	$hr = $configLine_array[2]
	$fl = $configLine_array[3]
	$fc = $configLine_array[4]
	$fr = $configLine_array[5]

	$filename = $filenameLine	

	if ( $filename -notmatch "^[A-Za-z]:\\" ) {
		$filename = (Get-Location).Path + '\' + $filename
	}

	$wb = $xl.Workbooks.Open($filename)
	
	try {
		# Replace when *From value exists.
		foreach ( $sheet in $wb.Worksheets ) {
			#FIXME: These six sequencial replacement is very boilerplate code.
			$hlOrig = $sheet.PageSetup.LeftHeader.ToString()
			if ( $hlOrig -ne $hl ) {
				if ( $dryrun ) {
					Write-Host "DRYRUN:" $sheet.name "Sheet LeftHeader change to " "$hlOrig to $hl" 
				} else {
					Write-Host $sheet.name "Sheet LeftHeader change to " "$hlOrig to $hl" 
					$sheet.PageSetup.LeftHeader = $hl
				} 
			}
			$hcOrig = $sheet.PageSetup.CenterHeader.ToString()
			if ( $hcOrig -ne $hc ) {
				if ( $dryrun ) {
					Write-Host "DRYRUN:" $sheet.name "Sheet CenterHeader change to " "$hcOrig to $hc" 
				} else {
					Write-Host $sheet.name "Sheet CenterHeader change to " "$hcOrig to $hc" 
					$sheet.PageSetup.CenterHeader = $hc
				}
			}
			$hrOrig = $sheet.PageSetup.RightHeader.ToString()
			if ( $hrOrig -ne $hr ) {
				if ( $dryrun ) {
					Write-Host "DRYRUN:" $sheet.name "Sheet RightHeader change to " "$hrOrig to $hr" 
				} else {
					Write-Host $sheet.name "Sheet RightHeader change to " "$hrOrig to $hr" 
					$sheet.PageSetup.RightHeader = $hr
				}
			}
			$flOrig = $sheet.PageSetup.LeftFooter.ToString()
			if ( $flOrig -ne $fl ) {
				if ( $dryrun ) {
					Write-Host "DRYRUN:" $sheet.name "Sheet LeftFooter change to " "$flOrig to $fl" 
				} else {
					Write-Host $sheet.name "Sheet LeftFooter change to " "$flOrig to $fl" 
					$sheet.PageSetup.LeftFooter = $fl
				}
			}
			$fcOrig = $sheet.PageSetup.CenterFooter.ToString()
			if ( $fcOrig -ne $fc ) {
				if ( $dryrun ) {
					Write-Host "DRYRUN:" $sheet.name "Sheet CenterFooter change to " "$fcOrig to $fc" 
				} else {
					Write-Host $sheet.name "Sheet CenterFooter change to " "$fcOrig to $fc" 
					$sheet.PageSetup.CenterFooter = $fc
				}
			}
			$frOrig = $sheet.PageSetup.RightFooter.ToString()
			if ( $frOrig -ne $fr ) {
				if ( $dryrun ) {
					Write-Host "DRYRUN:" $sheet.name "Sheet RightFooter change to " "$frOrig to $fr" 
				} else {
					Write-Host $sheet.name "Sheet RightFooter change to " "$frOrig to $fr" 
					$sheet.PageSetup.RightFooter = $fr
				}
			}
		}
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
		echo "DRY-RUN mode"
		$dryrunBool = $True
	} else {
		$dryrunBool = $False
	}

	$repListFile = (Get-Content $ReplaceHFFileList) -as [string[]]
	
	$configed = $False
	foreach ( $line in $repListFile ) {
		# skip comment line starting #
		if ( $line -match "^#.*" ) {
			continue
		} 
		if ( -! $configed ) {
			# First valid TSV line parsed as header-left,header-center,header-right,footer-left,footer-center,foot-right
			echo "$line"
			$configLine = $line
			$configed = $True
		} else {
			Replace-HeaderFooter $xl $configLine $line -dryrun $dryrunBool
		}
	}
} catch [Exception] {
	echo $error[0].exception        
} finally {
	$xl.Quit()
	[System.Runtime.Interopservices.Marshal]::ReleaseComObject($xl)
}

# vim: tw=0 tabstop=4:
