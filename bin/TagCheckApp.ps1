#Script date: 20-Oct-2023
#Author: Israel Rojas Valera

#region Execution Time
$TimeQuery = ConvertFrom-AFRelativeTime -RelativeTime "*"
$TimeQuery = $TimeQuery.ToLocalTime() 
#endregion

#region Directories
$CurrentBinPath = $PSScriptRoot
$CurrentLogPath = (get-item $CurrentBinPath).parent.FullName + "\Log"
$CurrentLogFilePath = $CurrentLogPath + "\TagCheckApp_Log.csv"
$CurrentConfigPath = (get-item $CurrentBinPath).parent.FullName + "\Config"
$CurrentOutputPath = (get-item $CurrentBinPath).parent.FullName + "\Output"
$CurrentKPIPath = (get-item $CurrentBinPath).parent.FullName + "\Output\KPI"
$CurrentConfigFile = (get-item $CurrentBinPath).parent.FullName + "\Config\parameters.xml"
#endregion

(Get-Date).ToString("MM/dd/yyyy HH:mm:ss") + "," + "Inicio Ejecución" | Out-File $CurrentLogFilePath -Append

#region Config read
[xml]$xmlConfig = Get-Content -Path $CurrentConfigFile
$XML_AFServer = $xmlConfig.Config.AFConfig.AFServer
$XML_AFUser = $xmlConfig.Config.AFConfig.AFUser
$XML_AFPassword = $xmlConfig.Config.AFConfig.AFPassword
$XML_AFDatabase = $xmlConfig.Config.AFConfig.AFDatabase
#endregion

function CheckPITag{
    param(
        [String]$Str_PIServer,
        [String]$Str_InputCSV,
        [String]$Str_OutputCSV,
        [String]$Str_KPICSV,
        [String]$Str_AFElement
        )        

        #Query DA Time
        $QueryDATime = ConvertFrom-AFRelativeTime -RelativeTime "*"
        $QueryDATime = $QueryDATime.ToLocalTime()

        #Global Variables
        $N_Tags = 0
        $N_Tags_Error = 0

        try{
            #Write-Host "Log Path: " $CurrentLogFilePath
            $PIDA_Connection = Connect-PIDataArchive $Str_PIServer
			
            #(Get-Date).ToString("MM/dd/yyyy HH:mm:ss") + "," + "Procesamiento CSV: " + $Str_InputCSV | Out-File $CurrentLogFilePath -Append
			
            $CSVTagList=Import-CSV -path $Str_InputCSV -Delimiter ";"

            $CSVTagList | ForEach-Object {
                $TagName = $_.'Tag Name'
                $N_Tags = $N_Tags + 1
                $Minimum = $_.'Minimum'
                $Maximum = $_.'Maximum'
                $DataType = $_.'Type'
                $StaleTimeSeconds = $_.'StaleTimeSeconds'
                $PIPoint = Get-PIPoint -Name $TagName -Connection $PIDA_Connection
                $TimeLimit = (Get-Date).AddSeconds(-1*$StaleTimeSeconds)
                
                If ($PIPoint -ne $null){
                    try{
         
                        $TipoTag = $PIPoint.Point.Type.ToString()

                        $PISnapshotValue = Get-PIValue -PIPoint $PIPoint -Time (ConvertFrom-AFRelativeTime -RelativeTime "*") -ArchiveMode Previous
                        $LastArcVal_TS = $PISnapshotValue.TimeStamp                    
                        $LastArcVal_TS = $LastArcVal_TS.ToLocalTime()				
                        $LastArcVal_Val = $PISnapshotValue.Value
                        $LastArcVal_Status = [boolean]$PISnapshotValue.IsGood        
                              
                        If ($LastArcVal_Status) {
                            If ($LastArcVal_TS -lt $TimeLimit) {
                                #Stale
							    If ($N_Tags_Error -eq 0){
								    "TagName,TimeStamp,Value,Comment" | Out-File $Str_OutputCSV -Append
							    }
                                $N_Tags_Error = $N_Tags_Error + 1
                                $TagName + "," +  $LastArcVal_TS + "," + $LastArcVal_Val + ",Variable Congelada" | Out-File $Str_OutputCSV -Append
                            }else
                            {
                                If ($TipoTag.equals("Timestamp") -or $TipoTag.equals("String")){
                                    $FechaLimiteInferior = ConvertFrom-AFRelativeTime -RelativeTime $Minimum
                                    $FechaLimiteInferior = $FechaLimiteInferior.ToLocalTime()
                                    #Write-Host "FechaLimiteInferior: " + $FechaLimiteInferior
                                    $FechaLimiteSuperior = ConvertFrom-AFRelativeTime -RelativeTime $Maximum
                                    $FechaLimiteSuperior = $FechaLimiteSuperior.ToLocalTime()
                                    #Write-Host "FechaLimiteSuperior: " + $FechaLimiteSuperior                                 

                                    # Parse the string to a DateTime object
                                    $LastArcVal_Val_Time = [DateTime]::ParseExact($LastArcVal_Val, "yyyy-MM-ddTHH:mm:ss.fffZ", $null)

                                    If (($LastArcVal_Val_Time -lt $FechaLimiteInferior) -or ($LastArcVal_Val_Time -gt $FechaLimiteSuperior)) {
                                        #Limites
									    If ($N_Tags_Error -eq 0){
									        "TagName,TimeStamp,Value,Comment" | Out-File $Str_OutputCSV -Append
									    }
                                        $N_Tags_Error = $N_Tags_Error + 1
                                        $TagName + "," +  $LastArcVal_TS + "," + $LastArcVal_Val_Time + ",Valor fuera de Limites" | Out-File $Str_OutputCSV -Append
                                    }
                                }else{
                                    If ($LastArcVal_Val.ToString().Length -lt 10) {
                                        If (($LastArcVal_Val -lt $Minimum) -or ($LastArcVal_Val -gt $Maximum)) {
                                            #Limites
									        If ($N_Tags_Error -eq 0){
										        "TagName,TimeStamp,Value,Comment" | Out-File $Str_OutputCSV -Append
									        }
                                            $N_Tags_Error = $N_Tags_Error + 1                                        
                                            $TagName + "," +  $LastArcVal_TS + "," + $LastArcVal_Val + ",Valor fuera de Limites" | Out-File $Str_OutputCSV -Append
                                        }
                                    }else{

                                        If($LastArcVal_Val.ToString().SubString(0,5) -ne "State") {       
                                            If (($LastArcVal_Val -lt $Minimum) -or ($LastArcVal_Val -gt $Maximum)) {
                                                #Limites
										        If ($N_Tags_Error -eq 0){
											        "TagName,TimeStamp,Value,Comment" | Out-File $Str_OutputCSV -Append
										        }
                                                $N_Tags_Error = $N_Tags_Error + 1
                                                $TagName + "," +  $LastArcVal_TS + "," + $LastArcVal_Val + ",Valor fuera de Limites" | Out-File $Str_OutputCSV -Append
                                            }
                                        }else{
                                            $stateID = $LastArcVal_Val.StateSet
                                            If ($stateID -eq 0){ #It is a System Digital State (Error)
                                                If ($N_Tags_Error -eq 0){
											        "TagName,TimeStamp,Value,Comment" | Out-File $Str_OutputCSV -Append
										        }
										        $N_Tags_Error = $N_Tags_Error + 1
                                                $stateValue = $LastArcVal_Val.State
                                                $digitalState = Get-PIDigitalStateSet -ID $stateID -Connection $PIDA_Connection
                                                $resultState = $digitalState[$stateValue]
                                                $TagName + "," +  $LastArcVal_TS + "," + $LastArcVal_Val + ",Calidad en Error" | Out-File $Str_OutputCSV -Append
                                            }                        
                                        }
                                    }
                                }
                            
                            
                            }
                        }else{
                            #Ditital State
                            $stateID = $LastArcVal_Val.StateSet
                            If ($stateID -eq 0){ #It is a System Digital State (Error)
                                If ($N_Tags_Error -eq 0){
								    "TagName,TimeStamp,Value,Comment" | Out-File $Str_OutputCSV -Append
							    }
							    $N_Tags_Error = $N_Tags_Error + 1
                                $stateValue = $LastArcVal_Val.State
                                $digitalState = Get-PIDigitalStateSet -ID $stateID -Connection $PIDA_Connection
                                $resultState = $digitalState[$stateValue]
                                $TagName + "," +  $LastArcVal_TS + "," + $resultState + ",Calidad en Error" | Out-File $Str_OutputCSV -Append
                            }
                        }
                    }
                    catch {
					    If ($N_Tags_Error -eq 0){
						    "TagName,TimeStamp,Value,Comment" | Out-File $Str_OutputCSV -Append
					    }
					    $N_Tags_Error = $N_Tags_Error + 1					
                        $e = $_.Exception
                        $msg = $e.Message
					    $TagName + "," + (Get-Date) + ",Error de lectura," + $msg | Out-File $Str_OutputCSV -Append
                    }    
                }else{
                    If ($N_Tags_Error -eq 0){
						"TagName,TimeStamp,Value,Comment" | Out-File $Str_OutputCSV -Append
					}
					$N_Tags_Error = $N_Tags_Error + 1                    
                    $TagName + "," + $QueryDATime.ToString("MM/dd/yyyy HH:mm:ss") + ",N/A,Punto PI No Existe" | Out-File $Str_OutputCSV -Append
                }               
            }    

            If ($N_Tags_Error -eq 0) {
                "TagName,TimeStamp,Value,Comment" | Out-File $Str_OutputCSV -Append 
                ',' + $QueryDATime.ToString("MM/dd/yyyy HH:mm:ss") + ",,Sin errores encontrados en lista de Tags" | Out-File $Str_OutputCSV -Append
                 $QueryDATime.ToString("MM/dd/yyyy HH:mm:ss") + "," + $Str_AFElement + "," + $N_Tags + "," + $N_Tags_Error + "," + $Str_OutputCSV| Out-File $Str_KPICSV -Append
            }else
            {
                $QueryDATime.ToString("MM/dd/yyyy HH:mm:ss") + "," + $Str_AFElement + "," + $N_Tags + "," + $N_Tags_Error + "," + $Str_OutputCSV| Out-File $Str_KPICSV -Append
            }
            
            $PIDA_Connection.Disconnect()
           
        }
        catch{
            Write-Host "An error occurred:"
            Write-Host $_.Exception 
            $e = $_.Exception
            $msg = $e.Message
			(Get-Date) + "," + "Excepción: " + $e | Out-File $CurrentLogFilePath -Append
			(Get-Date) + "," + "Mensaje: " + $msg | Out-File $CurrentLogFilePath -Append			
        }        
}

#region Clean Output Files
(Get-Date).ToString("MM/dd/yyyy HH:mm:ss") + "," + "Limpieza Archivos salida" | Out-File $CurrentLogFilePath -Append
$CurrentOutputPath_ = $CurrentOutputPath + "\*.csv"
$CurrentKPIPath_ = $CurrentKPIPath + "\*.csv"
Remove-Item -Path $CurrentOutputPath_ -Recurse -Force
Remove-Item -Path $CurrentKPIPath_ -Recurse -Force
#Clear-Content -Path $CurrentOutputPath -Filter "*.csv" -Force
#Clear-Content -Path $CurrentKPIPath -Filter "*.csv" -Force
#endregion

#region Tag list XML Reading
$xmlConfig.Config.TagLists.TagList | ForEach-Object {

    #Read PI Server
    $XML_PIDAServer = $_.PIDAServer
    
    #Read CSV File
    $XML_CSVInput = Join-Path -Path $CurrentConfigPath -ChildPath $_.Name

    #Read Output Result File
    $CSV_Output = Join-Path -Path $CurrentOutputPath -ChildPath $_.OutputResultFile

    #Read KPI Result File
    $CSV_KPI = Join-Path -Path $CurrentKPIPath -ChildPath $_.OutputResultFile

    #Read PI AF Output Results Element
    $AFOutputElement = $_.AFOutputElement
	
	#(Get-Date).ToString("MM/dd/yyyy HH:mm:ss") + "," + "Procesamiento CSV: " + $XML_CSVInput | Out-File $CurrentLogFilePath -Append
    Start-Job -ScriptBlock ${Function:CheckPITag} -ArgumentList $XML_PIDAServer, $XML_CSVInput, $CSV_Output, $CSV_KPI, $AFOutputElement, $global:array_results| Out-Null
    Start-Sleep -Seconds 1
}
#endregion

#region PI DA & PI AF Connection
        $secure_pass = ConvertTo-SecureString -String $XML_AFPassword -AsPlainText -Force
        $credentials = New-Object System.Management.Automation.PSCredential ($XML_AFUser, $secure_pass)
        try{
            $AFServer = Get-AFServer $XML_AFServer
            $AF_Connection = Connect-AFServer -WindowsCredential $credentials -AFServer $AFServer
            $AFDB = Get-AFDatabase -Name $XML_AFDatabase -AFServer $AFServer
            $AFRootElement = Get-AFElement -AFDatabase $AFDB -Name "Output"
        }catch { 
            $e = $_.Exception
            $msg = $e.Message
            while ($e.InnerException) {
                $e = $e.InnerException
                $msg += "`n" + $e.Message
                }
            $msg
        }
#endregion

#region Load Results to PI AF

Get-Job | Wait-Job| Receive-Job

$KPI_Files = Get-ChildItem $CurrentKPIPath

If ($AFRootElement -ne $null){
    #Limpieza de Tabla resultados
    $AF_Tables = $AFDB.Tables
    $AF_Tables["Results"].Table.Rows.Clear()
    foreach ($f in $KPI_Files){
        $KPI_CSV_Data = Import-CSV -path $f.FullName -Delimiter "," -Header 'Fecha', 'ElementoAF', 'TotalTags', 'TagsError', 'ArchivoDetalleTags'
        
        $Last_Record = $KPI_CSV_Data | Select-Object -Last 1
        
        $AFListaTag = Get-AFElement -AFElement $AFRootElement -Name $Last_Record.ElementoAF
        $AFTagCount = Get-AFAttribute -AFElement $AFListaTag -Name "Tag Count"
        ## %error
        $AFTagErrorPerc = Get-AFAttribute -AFElement $AFListaTag -Name "Porcentaje Error"
        ##
        $AFTagErrorCount = Get-AFAttribute -AFElement $AFListaTag -Name "Tag Error Count"
        $AFLastCheck = Get-AFAttribute -AFElement $AFListaTag -Name "LastCheck"
        $AFTagErrorList = Get-AFAttribute -AFElement $AFListaTag -Name "Tag Error List"
         
        ## calculo de % error
         $ErrorPerc= ($Last_Record.TagsError/$Last_Record.TotalTags)*100
        ## 
            
        #Update PI AF Attributes
        Set-Variable -Name AFLastCheck_AFV -Value (New-Object 'OSIsoft.AF.Asset.AFValue')
        $AFLastCheck_AFV.Value = $Last_Record.Fecha.ToString()
        $AFLastCheck.SetValue($AFLastCheck_AFV)
        Set-Variable -Name AFTagCount_AFV -Value (New-Object 'OSIsoft.AF.Asset.AFValue')
        $AFTagCount_AFV.Value = $Last_Record.TotalTags
        $AFTagCount.SetValue($AFTagCount_AFV)
        Set-Variable -Name AFTagErrorCount_AFV -Value (New-Object 'OSIsoft.AF.Asset.AFValue')
        $AFTagErrorCount_AFV.Value = $Last_Record.TagsError
        $AFTagErrorCount.SetValue($AFTagErrorCount_AFV)
        Set-Variable -Name AFTagErrorList_AFV -Value (New-Object 'OSIsoft.AF.Asset.AFFile')
        $AFTagErrorList_AFV.Upload($Last_Record.ArchivoDetalleTags)
        $AFTagErrorList.SetValue($AFTagErrorList_AFV)
        ## set tag % error
        Set-Variable -Name AFTagErrorPerc_AFV -Value (New-Object 'OSIsoft.AF.Asset.AFValue')
        $AFTagErrorPerc_AFV.Value = $ErrorPerc
        $AFTagErrorPerc_AFV.TimeStamp = Get-Date
        $AFTagErrorPerc.SetValue($AFTagErrorPerc_AFV)
        ##

        #Update Results Table
        $rownumber = 0
        $CSV_Output_Tags = Import-CSV -path $Last_Record.ArchivoDetalleTags -Delimiter "," -Header 'TagName', 'TimeStamp', 'Value', 'Comment'
        $CSV_Output_Tags | ForEach-Object {
            If ($rownumber -ne 0){
                $Output_Tag = $_.'TagNAme'
                $Output_Fecha = $_.'TimeStamp'                
                $Output_Fecha = [DateTime]::ParseExact($Output_Fecha, "MM/dd/yyyy HH:mm:ss", $null)
                $Output_Fecha = $Output_Fecha.ToString("yyyy-MM-dd HH:mm:ss")
                $Output_Valor = $_.'Value'
                $Output_Observacion = $_.'Comment'
                
                $AF_Tables["Results"].Table.Rows.Add($Last_Record.ElementoAF,$Output_Tag,$Output_Fecha,$Output_Valor,$Output_Observacion)
            }
            $rownumber = $rownumber +1
        }
    }
    $AF_Checkin = New-AFCheckIn($AFDB)
}

    #endregion

    #region Copy Results
    (Get-Date).ToString("MM/dd/yyyy HH:mm:ss") + "," + "Copia de Resultados a PI Vision" | Out-File $CurrentLogFilePath -Append
    $Source = "C:\TagCheck_App\V1\Output"
    $Destination = "\\192.168.254.157\Output_TagCheckApp\"
    Copy-Item -Path $Source\*.csv -Destination $Destination -Force
    #endregion

    #region Execution Time
    $TimeQuery_End = ConvertFrom-AFRelativeTime -RelativeTime "*"
    $TimeQuery_End = $TimeQuery.ToLocalTime() 
#endregion
   (Get-Date).ToString("MM/dd/yyyy HH:mm:ss") + "," + "Término Ejecución" | Out-File $CurrentLogFilePath -Append