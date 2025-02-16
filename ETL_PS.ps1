param (        
    [Parameter(Mandatory=$false)][string]$csvDir,
    [Parameter(Mandatory=$false)][string]$customTime,     
    [Parameter(Mandatory=$false)][double]$CNY_EXCHANGE_RATE, 
    [Parameter(Mandatory=$false)][double]$EUR_EXCHANGE_RATE
)

class CustomUtility {
    static [string]$LOGS_FILE_PATH = ".\logs\logs.log"
    static [string]$CUSTOM_DATE_TIME = '2025-01-01 02:00:00'
    static [string]$CSV_DIR = ".\csv_sales\"

    static [void] Log([string]$eventToLog){
        $currentDateTime = Get-Date
        $eventToLog = [string]$currentDateTime + " " +$eventToLog
        Add-Content -Path ([CustomUtility]::LOGS_FILE_PATH) -Value $eventToLog
    }
}

[double]$global:CNY_EXCHANGE_RATE = 0.13
[double]$global:EUR_EXCHANGE_RATE = 1.039

# ========================= PARAMETER CHECKS ========================= 

if([string]::IsNullOrEmpty($csvDir)){ 
    Write-Output "CSV directory is empty, defaulting to $([CustomUtility]::CSV_DIR)"
} else {
    [CustomUtility]::CSV_DIR = $csvDir
    Write-Output "CSV directory: $([CustomUtility]::CSV_DIR)"
}

if([string]::IsNullOrEmpty($customTime)){ 
    Write-Output "Custom time is empty, defaulting to $([CustomUtility]::CUSTOM_DATE_TIME)"
} else {
    [CustomUtility]::CUSTOM_DATE_TIME = $customTime
    Write-Output "Custom time : $([CustomUtility]::CUSTOM_DATE_TIME)"
}

if($CNY_EXCHANGE_RATE -eq 0){ 
    Write-Output "CNY exchange rate is empty, defaulting to $($global:CNY_EXCHANGE_RATE)"
} else {
    $global:CNY_EXCHANGE_RATE = $CNY_EXCHANGE_RATE
    Write-Output "CNY exchange rate: $($global:CNY_EXCHANGE_RATE)"
}

if($EUR_EXCHANGE_RATE -eq 0){ 
    Write-Output "EUR exchange rate is empty, defaulting to $($global:EUR_EXCHANGE_RATE)"
} else {
    $global:EUR_EXCHANGE_RATE = $EUR_EXCHANGE_RATE
    Write-Output "EUR exchange rate: $($global:EUR_EXCHANGE_RATE)"
}

# =========================  BASE CLASSES/PARENTS ========================= 

class DataSource {
    [string]$path 

    DataSource([string]$path){
        $this.path = $path
    }  
}

class DataTransformer {
    [string]$path 

    DataTransformer([string]$path){
        $this.path = $path
    }

    [System.Collections.ArrayList]Transform(){
        return @()
    }
}

class DataLoader {
    [string]$path 

    DataLoader([string]$path){
        $this.path = $path
    }    
}

# =========================  INHERITORS/CHILDREN ========================= 

class CSVDataSource : DataSource{
    [System.Collections.ArrayList]$FOUND_CSV_FILES
    
    CSVDataSource ([string]$salesCSVDir) : base($path) {
        if(Test-Path -Path $salesCSVDir){
            $this.FOUND_CSV_FILES = @(Get-ChildItem $salesCSVDir -Filter *.csv)
        } else {
            [CustomUtility]::Log("[ERROR] CSV path not found! Exitting...")
            exit
        }        
    }

    [bool] SearchForCSV(){
        if ($this.FOUND_CSV_FILES.Length -eq 0){            
            [CustomUtility]::Log("[WARN] No csv files at the moment. Exitting...")       
            exit              
        }
        [CustomUtility]::Log("[INFO] CSV files found!")        
        return $true
    }    

    [System.Collections.ArrayList] ExtractData(){        
        [System.Collections.ArrayList]$csvArr = [System.Collections.ArrayList]::new()
        [string]$path = ""

        # Load csv into variable/array
        foreach($item in $this.FOUND_CSV_FILES){                                    
            $path = [CustomUtility]::CSV_DIR + $item.Name
            try{
                $csv = Import-Csv -Path $path
                $csvArr.Add($csv)
                [CustomUtility]::Log("[INFO] Extracted CSV: " + $csv)
            } catch{
                [CustomUtility]::Log("[ERROR] Failed to extract data from $($path)")
            }                                                                 
        }                        
        return $csvArr
    }             
}

class CSVDataTransformer : DataTransformer{
    $filesPath
    [double]$exchangeRate

    CSVDataTransformer([string]$filesPath) : base($path) {
        $this.filesPath = $filesPath
    }
    
    [double] ConvertToUSD([double]$amount, [double]$exchangeRate){    
        return $amount * $exchangeRate    
    }

    # Standardize currency into USD and unify into one format: mm-dd-yyyy
    [System.Collections.ArrayList]Transform([System.Collections.ArrayList] $csvList){                

        # Transform currency and datetime
        Foreach ($csvFile in $csvList){
            Foreach ($csvRow in $csvFile){
                $csvRow.Amount = $csvRow.Amount -replace "\s", ""

                if($csvRow.Currency -eq "EUR"){
                    $this.exchangeRate = $global:EUR_EXCHANGE_RATE
                } elseif($csvRow.Currency -eq "CNY"){
                    $this.exchangeRate = $global:CNY_EXCHANGE_RATE
                } else {
                    $this.exchangeRate = 1.00
                }

                $csvRow.Amount = $this.ConvertToUSD($csvRow.Amount, $this.exchangeRate)
                $csvRow.Currency = "USD"

                # Reformat date
                $newDate = Get-Date $csvRow.Date
                $csvRow.Date = "$($newDate.Month)-$($newDate.Day)-$($newDate.Year)"                   
            }
        }
        return $csvList
    }        
}

class CSVDataLoader : DataLoader
{     
    [string]$excelReportPath

    # Accept only the final excel report
    CSVDataLoader($excelReportPath) : base($path){
        $this.excelReportPath = $excelReportPath
    }

    [void] GenerateFinalExcelReport([System.Collections.ArrayList]$csvList){ 
        $excel = New-Object -ComObject Excel.Application
        $excel.Visible = $false
        
        $workbook = $excel.Workbooks.Add()
        $worksheet = $workbook.Worksheets.Item(1)
        
        # Write header
        $worksheet.Cells.Item(1, 1) = "Country"
        $worksheet.Cells.Item(1, 2) = "Amount"
        $worksheet.Cells.Item(1, 3) = "Currency"
        $worksheet.Cells.Item(1, 4) = "Date"

        $rowIdx = 2
        Foreach ($csv in $csvList){                 
            Foreach ($csvRow in $csv){            
                # (row, col)
                $worksheet.Cells.Item($rowIdx, 1) = $csvRow.Country
                $worksheet.Cells.Item($rowIdx, 2) = $csvRow.Amount
                $worksheet.Cells.Item($rowIdx, 3) = $csvRow.Currency
                $worksheet.Cells.Item($rowIdx, 4) = $csvRow.Date        
                $rowIdx += 1
            }
        }
        
        # Save and close
        try{            
            $scriptPath = $PSScriptRoot
            $excelPath = Join-Path $scriptPath "final_report.xlsx"
            $workbook.SaveAs($excelPath)
            $workbook.Close()
            $excel.Quit()
        } catch{
            [CustomUtility]::Log("[ERROR] Failed to generate final excel report: $_")
            throw
        } finally {
            # Clean up COM objects
            if ($workbook) { [System.Runtime.Interopservices.Marshal]::ReleaseComObject($workbook) }
            if ($excel) { [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) }
            [System.GC]::Collect()
        }        
    }
}

# MAIN INTERFACE
class ETLPipeline {
    [CSVDataSource]$sourceClass
    [CSVDataTransformer]$transformerClass
    [CSVDataLoader]$loaderClass
    
    [datetime]$CURRENT_TIME
    [datetime]$TIME_TO_PROCESS_FILES
    [string]$LOG_FILE_PATH

    [System.Collections.ArrayList]$transformedCSVList

    [void] BuildLog([string]$logFilePath){        
        if(-not (Test-Path $logFilePath)){
            New-Item $logFilePath
        }    
    }
    
    [bool] ShouldExecute($currentTime, $timeToProcess, $logFilePath){
        if($currentTime.TimeOfDay -eq $timeToProcess.TimeOfDay){
            Write-Output "Time to process"
            return $true
        } else {
            Write-Output "Don't process yet"
            Log -logFilepath $logFilePath -eventToLog "[INFO] Processing time not reached."
            return $false
        }
    }
    
    ETLPipeline([string]$sourcePath, [string]$logFilePath, [string]$CUSTOM_DATE_TIME) {
        $this.CURRENT_TIME = Get-Date
        $this.TIME_TO_PROCESS_FILES = Get-Date $CUSTOM_DATE_TIME
        $this.LOG_FILE_PATH = $logFilePath        
    }

    [void] Execute(){
        # Extract
        $this.BuildLog([CustomUtility]::LOGS_FILE_PATH)        
        $dataSourceOBJ = New-Object CSVDataSource -ArgumentList ([CustomUtility]::CSV_DIR)
        $dataSourceOBJ.SearchForCSV()
        $importedCSVFiles = $dataSourceOBJ.ExtractData()
        
        # Transform
        $dataTransformerOBJ = New-Object CSVDataTransformer -ArgumentList ([CustomUtility]::CSV_DIR)
        $this.transformedCSVList = $dataTransformerOBJ.Transform($importedCSVFiles)

        # Load
        $dataLoaderOBJ = New-Object CSVDataLoader -ArgumentList ".\"        
        $dataLoaderOBJ.GenerateFinalExcelReport($this.transformedCSVList)

    }
}

$ETL = New-Object ETLPipeline -ArgumentList [CustomUtility]::CSV_DIR, [CustomUtility]::logsFilePath, ([CustomUtility]::CUSTOM_DATE_TIME)
$ETL.Execute()