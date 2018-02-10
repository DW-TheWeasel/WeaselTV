<#
WEASEL TV channel generator
Required: https://dev.mysql.com/downloads/connector/net/
#>

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

function Get-MediaLength ([string]$path) {
    # Returns media lenth (duration) in seconds
    # Cleanup filepath for windows
    $path = $path.Replace('smb:', '')
    $path = $path.Replace('/', '\')
    # Get Length
    $Folder = (get-item $path).Directory.FullName
    $File = (get-item $path).Name
    $LengthColumn = 27
    $objShell = New-Object -ComObject Shell.Application 
    $objFolder = $objShell.Namespace($Folder)
    $objFile = $objFolder.ParseName($File)
    $Length = $objFolder.GetDetailsOf($objFile, $LengthColumn)
    $LengthInSec = [TimeSpan]::Parse($Length).TotalSeconds
    $objShell.Dispose
    return $LengthInSec
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
$sQueryEpisodes = 'SELECT c09, c12, c13, c00, c01, c18  FROM episode WHERE episode.idShow = ' # C09="Episode length in minutes (depricated, need to get media info directly)"  C12="Season Number"  C13="Episode Number"  C00="Episode Title" C01="Plot Summary"  C18="Path to episode file" 
$sQueryGenreCh = "SELECT 
currentgenre, COUNT(*)
FROM
myvideos107.movie t1
    JOIN
(SELECT DISTINCT
    genre.name AS currentgenre
FROM
    myvideos107.genre) t2 ON t1.c14 LIKE CONCAT('%', currentgenre, '%')
WHERE
t1.c22 NOT LIKE '%/Media/Video/Ad%'
GROUP BY currentgenre
ORDER BY COUNT(*) DESC
LIMIT 15;" # Gets top 15 genres found in movie.c14 genre string (this is what smartplaylists use, so we need to use it instead of the genre table)
$sQueryMoviesInGenre = "SELECT 
*
FROM
    myvideos107.movie
WHERE
    myvideos107.movie.c22 NOT LIKE '%/Media/Video/Ad%'
        AND myvideos107.movie.c14 LIKE '%???%'
;"


# Query data
$aTvNetworks = Invoke-MySQLQuery $conn $sQueryNetworks      # TV Networks
$aTvIndex = Invoke-MySQLQuery $conn $sQueryIndex            # TV Index
$aGenreList = Invoke-MySQLQuery $conn $sQueryGenreCh        # Top movie genres


#Parse genre movie channels
$aGenreCh = @{}
foreach ($genre in $aGenreList) {
    if ([string]::IsNullOrEmpty($genre.currentgenre)) {continue}		                              # Skip blanks (first entry from query results always seems to be blank)
    $aTemp = Invoke-MySQLQuery $conn $sQueryMoviesInGenre.Replace('???',$genre.currentgenre)          # Temp array/query for all movies in a given genre from $aGenreList
    $aGenreCh.Add($genre.currentgenre, $aTemp)                                                        # Add genre as key and array of movies for value
}

# Parse TV Networks, Network Episodes
$aNetworkEpisodes = @{}                                                                               # TV Show Episodes by Network hashtable array (.c001="Show Name", .c00="Episode Name")
foreach ($network in $aTvNetworks) {
    if ([string]::IsNullOrEmpty($network.C14)) {continue}		                                      # Skip blanks (first entry from query results always seems to be blank)
    $sTemp = $sQueryNetworkEps + "'" + $network.C14 + "'"                                             # Build temp query for all TV Episodes for single Network by network.C14
    $aTemp = Invoke-MySQLQuery $conn $sTemp                                                           # Temp array/query for all TV Episodes for single Network by network.C14
    $aNetworkEpisodes.Add($network.C14, $aTemp)                                                       # Add network as key and array of episodes for value
}

# Parse TV Shows, TV Index, Show Episodes
$aTvShowEpisodes = @{}                                                                                  # TV Show Episodes by idShow hashtable array (key="Show Name", .c00="Episode Name")
foreach ($index in $aTvIndex) {
    if ([string]::IsNullOrEmpty($index.C00) -And [string]::IsNullOrEmpty($index.idShow)) {continue}		# Skip blanks (first entry from query results always seems to be blank)
    $sTemp = $sQueryEpisodes + $index.idShow + ";"                                                      # Build temp query for all TV Episodes for single show by idShow
    $aTemp = Invoke-MySQLQuery $conn $sTemp                                                             # Temp array/query for all TV Episodes for single show by idShow
    $aTvShowEpisodes.Add($index.C00, $aTemp)                                                            # Add TV show name as key and array of episodes for value
}

# Temp file management (remove old temp dirs and recreate blank structure)
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

# Build our settings2.xml string to later be written to file
$sS2XML = ""                                                              # Create openining XML
$sS2XML += "<settings>"
$sS2XML += "`n"
$sS2XML += "    <setting id=`"Version`" value=`"2.4.5`" />"
$sS2XML += "`n"

# Parse TV networks and enter channel entries to settings2.xml temp string
$iCount = 1
foreach ($network in $aTvNetworks) {
    if ([string]::IsNullOrEmpty($network.C14)) {continue}		          # Skip blanks (first entry from query results always seems to be blank)
    $sS2XML += ("    <setting id=`"Channel_" + $iCount + "_type`" value=`"1`" />")
    $sS2XML += "`n"
    $sS2XML += ("    <setting id=`"Channel_" + $iCount + "_1`" value=`"" + $network.C14 + "`" />")
    $sS2XML += "`n"
    $sS2XML += ("    <setting id=`"Channel_" + $iCount + "_changed`" value=`"False`" />")
    $sS2XML += "`n"
    $sS2XML += ("    <setting id=`"Channel_" + $iCount + "_time`" value=`"1`" />")
    $sS2XML += "`n"
    $sS2XML += ("    <setting id=`"Channel_" + $iCount + "_rulecount`" value=`"1`" />")
    $sS2XML += "`n"
    $sS2XML += ("    <setting id=`"Channel_" + $iCount + "_rule_1_id`" value=`"12`" />")
    $sS2XML += "`n"
    $iCount++
}

# Parse TV series and enter channel entries to settings2.xml temp string
$iCount = 100
foreach ($tvShow in $aTvIndex) {
    if ([string]::IsNullOrEmpty($tvShow.c00)) {continue}		          # Skip blanks (first entry from query results always seems to be blank)
    $sS2XML += ("    <setting id=`"Channel_" + $iCount + "_type`" value=`"6`" />")
    $sS2XML += "`n"
    $sS2XML += ("    <setting id=`"Channel_" + $iCount + "_1`" value=`"" + $tvShow.c00 + "`" />")
    $sS2XML += "`n"
    $sS2XML += ("    <setting id=`"Channel_" + $iCount + "_changed`" value=`"False`" />")
    $sS2XML += "`n"
    $sS2XML += ("    <setting id=`"Channel_" + $iCount + "_time`" value=`"1`" />")
    $sS2XML += "`n"
    $sS2XML += ("    <setting id=`"Channel_" + $iCount + "_rulecount`" value=`"1`" />")
    $sS2XML += "`n"
    $sS2XML += ("    <setting id=`"Channel_" + $iCount + "_rule_1_id`" value=`"12`" />")
    $sS2XML += "`n"
    $iCount++
}

# Parse movie genre channels and enter them in to settings2.xml temp string
$iCount = 900
foreach ($genreCh in $aGenreCh.Keys) {
    if ([string]::IsNullOrEmpty($genreCh)) {continue}		          # Skip blanks (first entry from query results always seems to be blank)
    $sS2XML += ("    <setting id=`"Channel_" + $iCount + "_type`" value=`"0`" />")
    $sS2XML += "`n"
    $sS2XML += ("    <setting id=`"Channel_" + $iCount + "_1`" value=`"special://profile/playlists/video/" + $genreCh + ".xsp`" />")
    $sS2XML += "`n"
    $sS2XML += ("    <setting id=`"Channel_" + $iCount + "_changed`" value=`"False`" />")
    $sS2XML += "`n"
    $sS2XML += ("    <setting id=`"Channel_" + $iCount + "_time`" value=`"1`" />")
    $sS2XML += "`n"
    $iCount++
}

#MOVIES BY YEAR NEXT


# Closing xml string
$sS2XML += "    <setting id=`"LastResetTime`" value=`"1495239376`" />"
$sS2XML += "`n"
$sS2XML += "    <setting id=`"LastExitTime`" value=`"1495239391`" />"
$sS2XML += "`n"
$sS2XML += "</settings>"

# Preview the settings2 string for testing
# Write-Host ($sS2XML | Out-String)

# Remove current settings2.xml if it exists.  It shouldn't, but double checking.
$sFile = ($env:TEMP + "\PTV\settings2.xml")
if (Test-Path -PathType Leaf $sFile) {
    # $sPath already exists
    Remove-Item $sFile -Force                                         # Delete $sFile
}

# Write settings2.xml to temp dir
$sS2XML | Out-File -FilePath $sFile -Force



















# Cleanup
Disconnect-MySQL($conn)