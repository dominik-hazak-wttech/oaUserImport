$path = Read-Host "Please provide path for data file"
$type = Read-Host "Please provide file type (csv, xls)"
switch($type.ToLower()){
    xls {
        $sheet = Read-Host "Please provide Sheet name (with no name default sheet will be selected)"
        $rowNum = Read-Host "Please starting row number (by default first row will be chosen)"
        if ($sheet -and $rowNum){
            $bulkData = Import-XLSX -Path $path -Sheet $sheet -RowStart $rowNum
        }
        elseif ($rowNum){
            $bulkData = Import-XLSX -Path $path -RowStart $rowNum
        }
        elseif ($sheet){
            $bulkData = Import-XLSX -Path $path -Sheet $sheet
        }
        else{
            $bulkData = Import-XLSX -Path $path
        }
    }
    csv {
        $delim = Read-Host "What character divides data?"
        $bulkData = Import-Csv -Path $path -Delimiter $delim
    }
    default{
        Write-Host "Unknown file type $type. Please try again"
    }
}