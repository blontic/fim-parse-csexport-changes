<# 
Original Script By Carol Wapsphere (http://www.wapshere.com/missmiis/using-powershell-to-parse-a-csexport-file)
Script Rewrite By Jorge de Almeida Pinto (https://jorgequestforknowledge.wordpress.com/2013/02/08/parsing-a-csexport-generated-xml-file-into-a-scoped-csv-file/)
Script Customised To Only Output Changes [ADD/DELETE] And Dynamically Create Column Headers By Clayton Brady

Takes an XML file created by CSEXPORT, and produces a CSV file more suitable for opening in Excel.
Supports both single-valued attributes and multi-valued attributes

	.EXAMPLE
	.\fim-parse-csexport-changes.ps1 -sourceXML C:\MAExport.xml -targetCSV C:\MAExport.csv
#>

Param(
    [Parameter(Mandatory = $true)]
    [string] $sourceXML,
    [string] $targetCSV
)

# Object Types Of Interest - User or Group or * For All
$objectTypes = @("*")

# Read The Source XML File
[System.Xml.XmlDocument] $xmlCSExportDoc = New-Object System.Xml.XmlDocument
$xmlCSExportDoc.load($sourceXML)

# Check If CSV File Already Exists
If (Test-Path $targetCSV) {	Remove-Item -Path $targetCSV -Force }

# Dynamically Build CSV Headers
Write-Host "Creating CSV Headers ..."
$csvHeaderColumns = New-Object System.Collections.Generic.List[System.String]
$i = 0
ForEach ($csObject In $xmlCSExportDoc."cs-objects"."cs-object") {
    $i = $i + 1
    Write-Host $i
    $attrName = $csObject.'unapplied-export'.'delta'.'attr'
    $attrRefName = $csObject.'unapplied-export'.'delta'.'dn-attr'
    $attr = $attrName, $attrRefName
    ForEach ($name in $attr.name) {
        If ($csvHeaderColumns -notcontains $name) { $csvHeaderColumns.Add($name) }
        write-host $i $csvHeaderColumns
    }
}
$csvHeadersSorted = $csvHeaderColumns | Sort-Object
$csvHeadersPrefix = @("dn", "object-type", "operation-type")
$csvHeaderColumns = $csvHeadersPrefix + $csvHeadersSorted

# Write The CSV Headers To The CSV File
$csvHeader = $null
ForEach ($csvHeaderColumn In $csvHeaderColumns) {
    If ($csvHeader -eq $null) {	$csvHeader = $csvHeaderColumn }
    Else { $csvHeader = $csvHeader + "," + $csvHeaderColumn	}
}
Add-Content $targetCSV $csvHeader

# Get The Information For The Scoped Objects
ForEach ($csObject In $xmlCSExportDoc."cs-objects"."cs-object") {
    If ($objectTypes -Contains $csObject."object-type" -Or $objectTypes -Contains "*") {
        Write-Host "Parsing " $csObject."cs-dn"
        $csObjectHashTable = @{}
        $csObjectHashTable.Add("dn", $csObject."cs-dn")
        $csObjectHashTable.Add("object-type", $csObject."object-type")
        $csObjectHashTable.Add("operation-type", $csObject."unapplied-export".delta.operation)

        # Add Operation
        If ($csObject."unapplied-export".delta.operation -eq "add") {
            ForEach ($csObjectAttribute In $csObject."unapplied-export".delta.attr) {
                If ($csObjectAttribute.multivalued -eq "false") { $csObjectHashTable.Add($csObjectAttribute.name, $csObjectAttribute.value) } 
                ElseIf ($csObjectAttribute.multivalued -eq "true" -And $csObjectAttribute.type -ne "binary") {
                    $multivaluedAttrValues = $null
                    If ($csObjectAttribute.value -ne "" -And $csObjectAttribute.value -ne $null) {
                        ForEach ($value in $csObjectAttribute.value) {
                            If ([string]::IsNullOrEmpty($multivaluedAttrValues)) { $multivaluedAttrValues = $value }
                            Else { $multivaluedAttrValues += ";" + $value }
                        }
                        $csObjectHashTable.Add($csObjectAttribute.name, $multivaluedAttrValues)
                    }
                    Else { $csObjectHashTable.Add($csObjectAttribute.name, "") }
                }
            }
            ForEach ($csObjectAttribute In $csObject.'unapplied-export'.delta.'dn-attr') {
                If ($csObjectAttribute.multivalued -eq "false") { $csObjectHashTable.Add($csObjectAttribute.name, $csObjectAttribute.'dn-value'.dn) }
                ElseIf ($csObjectAttribute.multivalued -eq "true" -And $csObjectAttribute.type -ne "binary") {
                    $multivaluedAttrValues = $null
                    If ([string]::IsNullOrEmpty($csObjectAttribute.'dn-value')) {
                        ForEach ($value in $csObjectAttribute.'dn-value') {
                            If ([string]::IsNullOrEmpty($multivaluedAttrValues)) { $multivaluedAttrValues = $value.dn }
                            Else { $multivaluedAttrValues += ";" + $value.dn }
                        }
                        $csObjectHashTable.Add($csObjectAttribute.name, $multivaluedAttrValues)
                    }
                    Else { $csObjectHashTable.Add($csObjectAttribute.name, "") }
                }
            }
        }

        # Update Operation
        ElseIf ($csObject."unapplied-export".delta.operation -eq "update") {
            ForEach ($csObjectAttribute In $csObject."unapplied-export".delta.attr) {
                $Values = $null
                $SortedValues = $null
                $SortedValues = $csObjectAttribute.value | Sort-Object -Property "Operation"
                ForEach ($value in $SortedValues) {
                    If ($value.operation -eq "delete") { $operation = "[Delete]: " }
                    Else { $operation = "[Add]: " }

                    If ($csObjectAttribute.value.count -lt 2 -and $csObjectAttribute.multivalued -eq "true" -and $csObjectAttribute.type -ne "binary") { $Values = $operation + $value."#text" }
                    ElseIf ($csObjectAttribute.value.count -lt 2) { $Values = $operation + $value }
                    Else {
                        If ([string]::IsNullOrEmpty($Values)) { $Values = $operation + $value."#text" }
                        Else { $Values += ";" + $operation + $value."#text" }
                    }
                }
                $csObjectHashTable.Add($csObjectAttribute.name, $Values)
            }
            ForEach ($csObjectAttribute In $csObject."unapplied-export".delta.'dn-attr') {
                $Values = $null
                $SortedValues = $null
                $SortedValues = $csObjectAttribute.'dn-value' | Sort-Object -Property "Operation"
                ForEach ($value in $SortedValues) {
                    If ($value.operation -eq "delete") {$operation = "[Delete]: "}
                    Else {$operation = "[Add]: "}

                    If ($SortedValues.count -lt 2 -and $csObjectAttribute.type -ne "binary") { $Values = $operation + $value.dn }
                    Else {
                        If ([string]::IsNullOrEmpty($Values)) { $Values = $operation + $value.dn }
                        Else { $Values += ";" + $operation + $value.dn }
                    }
                }
                $csObjectHashTable.Add($csObjectAttribute.name, $Values)
            }
        }

        # Delete Operation (Nothing needed here. CSV will just report this user as a delete)

        # Append to CSV file
        $csvLine = ""
        ForEach ($csvHeaderColumn in $csvHeaderColumns) {
            If ($csObjectHashTable.Contains($csvHeaderColumn)) {
                If ($csvLine -eq "") { $csvLine = "`"" + $csObjectHashTable.Item($csvHeaderColumn) + "`"" }
                Else { $csvLine += "," + "`"" + $csObjectHashTable.Item($csvHeaderColumn) + "`"" }
            }
            Else {
                If ($csvLine -eq "") { $csvLine = "," }
                Else { $csvLine += "," }
            }
        }
        Add-Content $targetCSV $csvline
    }
}