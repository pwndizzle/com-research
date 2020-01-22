########################################
# Recursive COM Method Search by @pwndizzle
#
# This script will recursively enumerate COM methods in order to dump all methods or find methods containing specific keywords. Be warned this is extremely hacky code and there are probably better ways to do this!
#
########################################
# OPTIONS
# Inpath - A file containing CLSIDs to scan
# Outpath - Scan results will be output to this file.
# Keywords - Search for methods containing this keyword
# Dumpall - Disabled by default. If true, keyword search will not be used and instead all properties will be dumped for the CLSIDs submitted.
# Depth - How deep to recurse. Note that many objects were found to support infinite/circular referencing.
########################################

$inpath = 'clsids.txt';
$outpath = 'output.txt';
$keywords = 'execute'; 
$dumpall = 0; 
$depth = 2;

foreach($cid in Get-Content $inpath) {

$cid
try{
$Obj = [System.Activator]::CreateInstance([Type]::GetTypeFromCLSID($cid));

function recur([string]$recpath) {
	
	$path = $recpath;
	$recobj = $Obj;
	if($path){
		foreach($p in $path.split(".")) { 
			$recobj = $recobj.$p;
		}
	}
	
	$recobj.PSObject.Methods | ForEach-Object {
		if ($dumpall){
			$cid + " - " + $path + "." + $_.Name >> $outpath;
		} else {
			foreach($keyword in $keywords) { 
				if($_.Name -like "*$keyword*") { 
				$cid + " - " + $path + "." + $_.Name >> $outpath;
				}
			}			
		}
	}

	$recobj.PSObject.Properties | ForEach-Object {
		if (!$path -And $_.Name -notlike "*Parent*"){
			$newpath = $_.Name;
			recur($newpath);
		} elseif (-not ([string]::IsNullOrEmpty($_.Name))){
			$name = $_.Name;
			if($path -notlike "*$name*" ){
				$newpath = $path+"."+$_.Name;
				if(($newpath.Split('.')).count-1 -lt $depth -And $newpath -notlike "*.Parent*" -And $newpath -notlike "*.Formula*" -And $newpath -notlike "*.MailEnvelope*"){
					recur($newpath);
				}
			}
		}
	}
	return
}
$p = "";
recur($p);
$res = [System.Runtime.Interopservices.Marshal]::ReleaseComObject($Obj);
Start-Sleep -s 1;
}catch{}
}
