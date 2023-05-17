function New-Table {
    Param(
        [Parameter (Mandatory = $true)][string] $TableName,
        [Parameter (Mandatory = $true)][object[]] $tblArray,
        [Parameter (Mandatory = $true)][object[]] $columnArray
    )
    #write-host $tblArray
  Process{    
    # Create Table object
    $table = New-Object system.Data.DataTable $TableName

    # Define and Add Columns to table
    $ColumnCount = $columnArray.count
    for($i=0; $i -lt $ColumnCount;$i++){
        $var = New-Object system.Data.DataColumn $columnArray[$i],([string])
        $table.columns.add($var)  
    }

    ForEach ($item in $tblArray ) 
    {
        # Create a row
        $row = $table.NewRow()
    
        # Enter data in the row
        for($i=0; $i -lt $ColumnCount;$i++){
            $data = $columnArray[$i]
            $columnObj = $table.Columns[$i]
            $row.$columnObj = $item.$data
        }
    
        # Add the row to the table
        $table.Rows.Add($row)
    }
    return $table}
}