# WEASEL TV channel generator
# 
# Required: https://dev.mysql.com/downloads/connector/net/

Clear-Host #start with a blank console

function Connect-MySQL([string]$user, [string]$pass, [string]$MySQLHost, [string]$port, [string]$database) { 
    # Load MySQL .NET Connector Objects 
    [void][system.reflection.Assembly]::LoadWithPartialName("MySql.Data") 
 
    # Open Connection 
    $connStr = "server=" + $MySQLHost + ";port=" + $port + ";uid=" + $user + ";pwd=" + $pass + ";database=" + $database + ";Pooling=FALSE" 
    try {
        $conn = New-Object MySql.Data.MySqlClient.MySqlConnection($connStr) 
        $conn.Open()
    }
    catch [System.Management.Automation.PSArgumentException] {
        Write-Host "Unable to connect to MySQL server, do you have the MySQL connector installed..?"
        Write-Host $_
        Exit
    }
    catch {
        Write-Host "Unable to connect to MySQL server..."
        Write-Host $_.Exception.GetType().FullName
        Write-Host $_.Exception.Message
        exit
    }
    Write-Host "Connected to MySQL database $MySQLHost\$database"

    return $conn 
}

function Disconnect-MySQL($conn) {
    $conn.Close()
}

function Invoke-MySQLNonQuery($conn, [string]$query) { 
    # MySQLNonQuery - Insert/Update/Delete query where no return data is required
    $command = $conn.CreateCommand()                  # Create command object
    $command.CommandText = $query                     # Load query into object
    $RowsInserted = $command.ExecuteNonQuery()        # Execute command
    $command.Dispose()                                # Dispose of command object
    if ($RowsInserted) { 
        return $RowInserted 
    }
    else { 
        return $false 
    } 
} 

function Invoke-MySQLQuery($connMySQL, [string]$query) { 
    # MySQLQuery - Used for normal queries that return multiple values. Results need to be received into MySqlDataReader object
    $cmd = New-Object MySql.Data.MySqlClient.MySqlCommand($query, $connMySQL)    # Create SQL command
    $dataAdapter = New-Object MySql.Data.MySqlClient.MySqlDataAdapter($cmd)      # Create data adapter from query command
    $dataSet = New-Object System.Data.DataSet                                    # Create dataset
    $dataAdapter.Fill($dataSet, "data")                                          # Fill dataset from data adapter, with name "data"              
    $cmd.Dispose()
    return $dataSet.Tables["data"]                                               # Returns an array of results
    # EXAMPLE:
    # $query = "SELECT * FROM subnets;"
    # $result = Invoke-MySQLQuery $query
    # Write-Host ("Found " + $result.rows.count + " rows...")
    # $result | Format-Table
}

function Invoke-MySQLScalar([string]$query) {
    # MySQLScalar - Select etc query where a single value of return data is expected
    $cmd = $SQLconn.CreateCommand()                                             # Create command object
    $cmd.CommandText = $query                                                   # Load query into object
    $cmd.ExecuteScalar()                                                        # Execute command
}


# Connection Variables 
$user = 'kodi' 
$pass = 'kodi' 
$database = 'MyVideos107' 
$MySQLHost = 'TORRENTS'
$port = '3306' 

# Connect to MySQL Database 
$conn = Connect-MySQL $user $pass $MySQLHost $port $database








# Queries
$sQueryNetworks = 'Select Distinct C14 From tvshow ORDER BY lower( C14 );'
$sQueryIndex = 'Select Distinct C00, tvshow.idShow From tvshow ORDER BY lower( C00 );'
$sQueryNetworkEps = 'SELECT episode.c09, episode.c12, episode.c13, episode.c00, episode.c01, episode.c18, tvshow.c00 FROM episode INNER JOIN tvshow ON episode.idShow=tvshow.idShow WHERE tvshow.c14 LIKE ' # tvshow.C00="Show Title"  tvshow.C14="Studio"
$sQueryEpisodes = 'SELECT c09, c12, c13, c00, c01, c18  FROM episode WHERE episode.idShow = ' # C09="Episode length in minutes (converted to sec in DB)"  C12="Season Number"  C13="Episode Number"  C00="Episode Title" C01="Plot Summary"  C18="Path to episode file" 


# Network
#$aTvNetworks = Invoke-MySQLQuery $conn $sQueryNetworks
#$aTvNetworks | Format-Table

# Query data
$aTvNetworks = Invoke-MySQLQuery $conn $sQueryNetworks     # TV Networks
$aTvIndex = Invoke-MySQLQuery $conn $sQueryIndex           # TV Index

# Parse TV Networks, Network Episodes
$aNetworkEpisodes = @{}                                                                               # TV Show Episodes by Network hashtable array (.c001="Show Name", .c00="Episode Name")
foreach ($network in $aTvNetworks) {
    if ([string]::IsNullOrEmpty($network.C14)) {continue}		                                          # Skip blanks (first entry from query results always seems to be blank)
    $sTemp = $sQueryNetworkEps + "'" + $network.C14 + "'"                                             # Build temp query for all TV Episodes for single Network by network.C14
    $aTemp = Invoke-MySQLQuery $conn $sTemp                                                           # Temp array/query for all TV Episodes for single Network by network.C14
    $aNetworkEpisodes.Add($network.C14, $aTemp)                                                       # Add network as key and array of episodes for value
    #Write-Host $network.C14
}

# Parse TV Shows, TV Index, Show Episodes
$aTvShowEpisodes = @{}                                                                                  # TV Show Episodes by idShow hashtable array (key="Show Name", .c00="Episode Name")
$aTvShows = @()                                                                                         # TV Shows array
foreach ($index in $aTvIndex) {
    if ([string]::IsNullOrEmpty($index.C00) -And [string]::IsNullOrEmpty($index.idShow)) {continue}		  # Skip blanks (first entry from query results always seems to be blank)
    $aTvShows += , $index.C00                                                                           # Add to TV shows array
    $sTemp = $sQueryEpisodes + $index.idShow + ";"                                                      # Build temp query for all TV Episodes for single show by idShow
    $aTemp = Invoke-MySQLQuery $conn $sTemp                                                             # Temp array/query for all TV Episodes for single show by idShow
    $aTvShowEpisodes.Add($index.C00, $aTemp)                                                            # Add TV show name as key and array of episodes for value
    #Write-Host ($index.C00 + " - " + $index.idShow)
}

#Write-Host ($aTvShowEpisodes | Out-String)
#Write-Host ($aTvShowEpisodes["Drawn Together"] | Out-String)

#Write-Host ($aNetworkEpisodes | Out-String)
#Write-Host ($aNetworkEpisodes["ABC (US)"] | Out-String)

$aPaths = @()                                                             # Array of paths to remove and recreate empty
$aPaths += , ($env:TEMP + "\PTV")
$aPaths += , ($env:TEMP + "\PTV\cache")
$aPaths += , ($env:TEMP + "\MoviePlaylists")

foreach ($sPath in $aPaths) {
    if (Test-Path -PathType Container $sPath) {
        # $sPath already exists
        Remove-Item $sPath -Force -Recurse                                # Delete $sPath
        New-Item -ItemType Directory -Force -Path $sPath | Out-Null       # Recreate $sPath
    }
    else {
        # $sPath doesn't exist
        New-Item -ItemType Directory -Force -Path $sPath | Out-Null       # Create $sPath
    }
}

$sFile = ($env:TEMP + "\PTV\settings2.xml")
if (Test-Path -PathType Leaf $sFile) {
    # $sPath already exists
    Remove-Item $sFile -Force                                         # Delete $sFile
}














Disconnect-MySQL($conn)













# Test query
#$query = "SELECT Distinct C22 FROM movie;"
#$result = Invoke-MySQLQuery $conn $query
#Write-Host ("Found " + $result.rows.count + " rows...")
#$result | Format-Table

#Write-Host ($result.gettype())

#$result.foreach({Write-Host($_.ItemArray)})

#Write-Host $result[1]

#$result.GetType()
#$result | Get-Help
#$result | Get-Command
#$result | Get-Member
#($result | Get-Member)[-1].Name


#foreach ($element in $result) {
#	$element
#}








