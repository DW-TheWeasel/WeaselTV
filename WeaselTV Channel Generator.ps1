<#
WEASEL TV channel generator
Alpha v0.1.1
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
    #$LengthInSec = [TimeSpan]::Parse($Length).TotalSeconds
    try {
        $LengthInSec = [TimeSpan]::Parse($Length).TotalSeconds
    }
    catch {
        $LengthInSec = 0
    }
    $objShell = $null
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
$sQueryNetworks = "SELECT DISTINCT
    myvideos107.tvshow.C14
FROM
    myvideos107.tvshow
ORDER BY LOWER(C14)
;"

$sQueryIndex = "SELECT DISTINCT
    C00, myvideos107.tvshow.idShow
FROM
    myvideos107.tvshow
ORDER BY LOWER(C00)
;"

$sQueryNetworkEps = "SELECT 
    *
FROM
    myvideos107.episode
        INNER JOIN
    myvideos107.tvshow ON myvideos107.episode.idShow = myvideos107.tvshow.idShow
WHERE
    myvideos107.tvshow.c14 LIKE '???'
;"                                                                                                      # tvshow.C00="Show Title"  tvshow.C14="Studio"

$sQueryEpisodes = "SELECT 
    *
FROM
    myvideos107.episode
WHERE
    myvideos107.episode.idShow = ???
;"                                                                                                      # C09="Episode length in minutes (we will populate this as cache)"  C12="Season Number"  C13="Episode Number"  C00="Episode Title" C01="Plot Summary"  C18="Path to episode file" 

$sQueryUpdateEpLength = "UPDATE
    myvideos107.episode 
SET 
    myvideos107.episode.c09 = ???LENGTH???
WHERE
    myvideos107.episode.idEpisode = ???ID???
;"                                                                                                      # Update c09 episode length (using this as cache)

$sQueryUpdateMovieLength = "UPDATE
myvideos107.movie 
SET 
myvideos107.movie.c11 = ???LENGTH???
WHERE
myvideos107.movie.idMovie = ???ID???
;"                                                                                                      # Update c11 movie length (using this as cache)

$sQueryGenreCh = "SELECT 
currentgenre, COUNT(*)
FROM
myvideos107.movie t1
    JOIN
(SELECT DISTINCT
    myvideos107.genre.name AS currentgenre
FROM
    myvideos107.genre) t2 ON t1.c14 LIKE CONCAT('%', currentgenre, '%')
WHERE
t1.c22 NOT LIKE '%/Media/Video/Ad%'
GROUP BY currentgenre
ORDER BY COUNT(*) DESC
LIMIT 15
;"                                                                                                      # Gets top 15 genres found in movie.c14 genre string (this is what smartplaylists use, so we need to use it instead of the genre table)

$sQueryMoviesInGenre = "SELECT 
    *
FROM
    myvideos107.movie
WHERE
    myvideos107.movie.c22 NOT LIKE '%/Media/Video/Ad%'
        AND myvideos107.movie.c14 LIKE '%???%'
;"

$sQueryMoviesInDecade = "SELECT 
    *
FROM
    myvideos107.movie
WHERE
    myvideos107.movie.c22 NOT LIKE '%/Media/Video/Ad%'
        AND myvideos107.movie.premiered >= 'YYYY1-01-01'
        AND myvideos107.movie.premiered < 'YYYY2-01-01'
;"

# Query data
$aTvNetworks = Invoke-MySQLQuery $conn $sQueryNetworks                                                  # TV Networks
$aTvIndex = Invoke-MySQLQuery $conn $sQueryIndex                                                        # TV Index
$aGenreList = Invoke-MySQLQuery $conn $sQueryGenreCh                                                    # Top movie genres
$aDecadeList = ('1970', '1980', '1990', '2000', '2010')                                                     # Decade movie channels



# Parse TV Networks, Network Episodes
$aNetworkEpisodes = @{}                                                                                 # TV Show Episodes by Network hashtable array (.c001="Show Name", .c00="Episode Name")
foreach ($network in $aTvNetworks) {
    if ([string]::IsNullOrEmpty($network.C14)) {continue}		                                        # Skip blanks
    $sTemp = $sQueryNetworkEps.Replace('???', $network.C14)                                             # Build temp query for all TV Episodes for single Network by network.C14
    $aTemp = Invoke-MySQLQuery $conn $sTemp                                                             # Temp array/query for all TV Episodes for single Network by network.C14
    $aNetworkEpisodes.Add($network.C14, $aTemp)                                                         # Add network as key and array of episodes for value
}

# Parse TV Shows, TV Index, Show Episodes
$aTvShowEpisodes = @{}                                                                                  # TV Show Episodes by idShow hashtable array (key="Show Name", .c00="Episode Name")
foreach ($index in $aTvIndex) {
    if ([string]::IsNullOrEmpty($index.C00) -And [string]::IsNullOrEmpty($index.idShow)) {continue}		# Skip blanks
    $sTemp = $sQueryEpisodes.Replace('???', $index.idShow)                                              # Build temp query for all TV Episodes for single show by idShow
    $aTemp = Invoke-MySQLQuery $conn $sTemp                                                             # Temp array/query for all TV Episodes for single show by idShow
    $aTvShowEpisodes.Add($index.C00, $aTemp)                                                            # Add TV show name as key and array of episodes for value
}

# Parse genre movie channels
$aGenreCh = @{}
foreach ($genre in $aGenreList) {
    if ([string]::IsNullOrEmpty($genre.currentgenre)) {continue}		                                # Skip blanks
    $sTemp = $sQueryMoviesInGenre.Replace('???', $genre.currentgenre)                                   # Build temp query for all movies in a given genre from $aGenreList
    $aTemp = Invoke-MySQLQuery $conn $sTemp                                                             # Temp array/query for all movies in a given genre from $aGenreList
    $aGenreCh.Add($genre.currentgenre, $aTemp)                                                          # Add genre as key and array of movies for value
}

# Parse decade movie channels
$aDecadeCh = @{}
foreach ($decade in $aDecadeList) {
    if ([string]::IsNullOrEmpty($decade)) {continue}		                                            # Skip blanks
    $sTemp = $sQueryMoviesInDecade.Replace('YYYY1', [int]$decade).Replace('YYYY2', [int]$decade + 10)   # Build temp query for all movies in a given decade
    $aTemp = Invoke-MySQLQuery $conn $sTemp                                                             # Temp array/query for all movies in a given decade
    $aDecadeCh.Add([string]$decade, $aTemp)                                                             # Add decade as key and array of movies for value
}

# Temp file management (remove old temp dirs and recreate blank structure)
$aPaths = @()                                                                                           # Array of paths to remove and recreate empty
$aPaths += , ($env:TEMP + "\PTV")
$aPaths += , ($env:TEMP + "\PTV\cache")
$aPaths += , ($env:TEMP + "\MoviePlaylists")
foreach ($sPath in $aPaths) {
    if (Test-Path -PathType Container $sPath) {
        # $sPath already exists
        Remove-Item $sPath -Force -Recurse                                                              # Delete $sPath
        New-Item -ItemType Directory -Force -Path $sPath | Out-Null                                     # Recreate $sPath
    }
    else {
        # $sPath doesn't exist
        New-Item -ItemType Directory -Force -Path $sPath | Out-Null                                     # Create $sPath
    }
}


# Create m3u array
$aM3u = @{}

# Create xsp array
$aXSP = @{}

# Running totals
$iTotalShows = 0
$iTotalNetworks = 0
$iTotalShowEps = 0
$iTotalNetworkEps = 0
$activity2 = ''
$activity3 = ''
Write-Progress -activity "Totals:"  -status ("Networks: " + $iTotalNetworks + " | Network Episodes: " + $iTotalNetworkEps + " | Shows: " + $iTotalShows + " | Episodes: " + $iTotalShowEps) -Id 1

# Build our settings2.xml string to later be written to file
$sS2XML = ""                                                                                            # Create openining XML
$sS2XML += "<settings>"
$sS2XML += "`n"
$sS2XML += "    <setting id=`"Version`" value=`"2.4.5`" />"
$sS2XML += "`n"

# Parse TV networks and enter channel entries to settings2.xml temp string
$iCount = 1
$iCurCh = $iCount
foreach ($network in $aNetworkEpisodes.GetEnumerator() | Sort-Object Name) {
    if ([string]::IsNullOrEmpty($network.key)) {continue}		                                        # Skip blanks
    $iTotalNetworks++
    # if ($iCount -lt 4) {$iCount++;continue} # skip for testing
    $sS2XML += ("    <setting id=`"Channel_" + $iCount + "_type`" value=`"1`" />")
    $sS2XML += "`n"
    $sS2XML += ("    <setting id=`"Channel_" + $iCount + "_1`" value=`"" + $network.key + "`" />")
    $sS2XML += "`n"
    $sS2XML += ("    <setting id=`"Channel_" + $iCount + "_changed`" value=`"False`" />")
    $sS2XML += "`n"
    $sS2XML += ("    <setting id=`"Channel_" + $iCount + "_time`" value=`"1`" />")
    $sS2XML += "`n"
    $sS2XML += ("    <setting id=`"Channel_" + $iCount + "_rulecount`" value=`"1`" />")
    $sS2XML += "`n"
    $sS2XML += ("    <setting id=`"Channel_" + $iCount + "_rule_1_id`" value=`"12`" />")
    $sS2XML += "`n"
    
    if (($network.Value.c001 | Get-Unique).count -gt 1) {
        $sortednetwork = $network.Value | Get-Random -Count ([int]::MaxValue)                           # Randomize network episodes if more than 1 show is present on the network
    }
    else {
        $sortednetwork = $network.Value                                                                 # Do not sort if only 1 show on the network
    }

    $activity2 = "Scanning TV network"
    Write-Progress -activity $activity2 -status $network.key -Id 2 -percentComplete ((($iCount - $iCurCh) / $aNetworkEpisodes.Count) * 100)

    $iSubCount = 0
    $sShows = ''
    $iTotalCount = $sortednetwork.Count
    $sw = [System.Diagnostics.Stopwatch]::StartNew()

    $stringBuilder = New-Object System.Text.StringBuilder
    foreach ($episode in $sortednetwork.GetEnumerator()) {
        if ([string]::IsNullOrEmpty($episode.c18)) {continue}
        $iTotalNetworkEps++
        if ($episode.c09 -eq 0) {
            $MediaLength = Get-MediaLength($episode.c18)
            if ($MediaLength -ne 0) {
                # Update cache in db at c09
                $sTempQuery = $sQueryUpdateEpLength.Replace("???LENGTH???", $MediaLength).Replace("???ID???", $episode.idEpisode)
                Invoke-MySQLNonQuery $conn $sTempQuery
            }
        }
        else {
            # Use cached media length/duration
            $MediaLength = $episode.c09
        }
        

        if ($episode.c12.Length -eq 1) {$sSeason = '0' + $episode.c12} else {$sSeason = $episode.c12}   # Format SS
        if ($episode.c13.Length -eq 1) {$sEpisode = '0' + $episode.c13} else {$sEpisode = $episode.c13} # Format EE

        $null = $stringBuilder.Append('#EXTINF:')                                                                        # #EXTINF:
        $null = $stringBuilder.Append($MediaLength)                                                                      # Media length/duration
        $null = $stringBuilder.Append(',')                                                                               # ,
        $null = $stringBuilder.Append($episode.c001.Replace("`t", " ").Replace("`n", " ").Replace("`r", " "))            # Show name
        $null = $stringBuilder.Append('//')                                                                              # //
        $null = $stringBuilder.Append(('S' + $sSeason + 'E' + $sEpisode))                                                # SxxExx
        $null = $stringBuilder.Append(' - ')                                                                             #  - 
        $null = $stringBuilder.Append($episode.c00.Replace("`t", " ").Replace("`n", " ").Replace("`r", " "))             # Episode name
        $null = $stringBuilder.Append('//')                                                                              # //
        $null = $stringBuilder.Append($episode.c01.Replace("`t", " ").Replace("`n", " ").Replace("`r", " "))             # Episode description
        $null = $stringBuilder.Append("`n")                                                                              # New line
        $null = $stringBuilder.Append($episode.c18)                                                                      # File with full path
        $null = $stringBuilder.Append("`n")                                                                              # New line
        $iSubCount++
        # Progress
        if (-Not($sShows -match $episode.c001.Replace("`t", " ").Replace("`n", " ").Replace("`r", " ").Replace("*", "").Replace("(", "").Replace(")", ""))) {
            $sShows = $sShows + $episode.c001.Replace("`t", " ").Replace("`n", " ").Replace("`r", " ").Replace("*", "").Replace("(", "").Replace(")", "") + " | "
        }
        if ($sw.Elapsed.TotalMilliseconds -ge 10) {
            $activity3 = ("Generating m3u (" + $iSubCount + "/" + $iTotalCount + ")for:")
            Write-Progress -activity "Totals:"  -status ("Networks: " + $iTotalNetworks + " | Network Episodes: " + $iTotalNetworkEps + " | Shows: " + $iTotalShows + " | Episodes: " + $iTotalShowEps) -Id 1
            Write-Progress -activity $activity3 -status $sShows -Id 3 -percentComplete (($iSubCount / $iTotalCount) * 100)
            $sw.Reset(); $sw.Start()
        }

    }
    $sw = $null

    # Add m3u data to $aM3u
    $sM3uEntry = $stringBuilder.ToString()
    $stringBuilder = $null
    $aM3u.Add('channel_' + $iCount + '.m3u', $sM3uEntry)
    $iCount++
}

# Parse TV series and enter channel entries to settings2.xml temp string
$iCount = 100
$iCurCh = $iCount
foreach ($tvShow in $aTvShowEpisodes.GetEnumerator() | Sort-Object Name) {
    if ([string]::IsNullOrEmpty($tvShow.key)) {continue}		                                        # Skip blanks
    $iTotalShows++
    $sS2XML += ("    <setting id=`"Channel_" + $iCount + "_type`" value=`"6`" />")
    $sS2XML += "`n"
    $sS2XML += ("    <setting id=`"Channel_" + $iCount + "_1`" value=`"" + $tvShow.key + "`" />")
    $sS2XML += "`n"
    $sS2XML += ("    <setting id=`"Channel_" + $iCount + "_changed`" value=`"False`" />")
    $sS2XML += "`n"
    $sS2XML += ("    <setting id=`"Channel_" + $iCount + "_time`" value=`"1`" />")
    $sS2XML += "`n"
    $sS2XML += ("    <setting id=`"Channel_" + $iCount + "_rulecount`" value=`"1`" />")
    $sS2XML += "`n"
    $sS2XML += ("    <setting id=`"Channel_" + $iCount + "_rule_1_id`" value=`"12`" />")
    $sS2XML += "`n"

    $activity2 = "Scanning TV shows"
    Write-Progress -activity $activity2 -status $tvShow.key -Id 2 -percentComplete ((($iCount - $iCurCh) / $aTvShowEpisodes.Count) * 100)

    $iSubCount = 0
    $sShows = $tvShow.key
    $iTotalCount = $tvShow.Value.Count
    $sw = [System.Diagnostics.Stopwatch]::StartNew()
    
    $stringBuilder = New-Object System.Text.StringBuilder
    foreach ($episode in $tvShow.Value.GetEnumerator()) {
        if ([string]::IsNullOrEmpty($episode.c18)) {continue}
        $iTotalShowEps++
        if ($episode.c12.Length -eq 1) {$sSeason = '0' + $episode.c12} else {$sSeason = $episode.c12}   # Format SS
        if ($episode.c13.Length -eq 1) {$sEpisode = '0' + $episode.c13} else {$sEpisode = $episode.c13} # Format EE
        # Check for cached media length/duration in c09. Cache here if we have to look it up
        if ($episode.c09 -eq 0) {
            $MediaLength = Get-MediaLength($episode.c18)
            if ($MediaLength -ne 0) {
                # Update cache in db at c09
                $sTempQuery = $sQueryUpdateEpLength.Replace("???LENGTH???", $MediaLength).Replace("???ID???", $episode.idEpisode)
                Invoke-MySQLNonQuery $conn $sTempQuery
            }
        }
        else {
            # Use cached media length/duration
            $MediaLength = $episode.c09
        }
        $null = $stringBuilder.Append('#EXTINF:')                                                                        # #EXTINF:
        $null = $stringBuilder.Append($MediaLength)                                                                      # Media length/duration
        $null = $stringBuilder.Append(',')                                                                               # ,
        $null = $stringBuilder.Append($episode.c00.Replace("`t", " ").Replace("`n", " ").Replace("`r", " "))             # Episode name
        $null = $stringBuilder.Append('//')                                                                              # //
        $null = $stringBuilder.Append(('S' + $sSeason + 'E' + $sEpisode))                                                # SxxExx
        $null = $stringBuilder.Append(' - ')                                                                             #  - 
        $null = $stringBuilder.Append($tvShow.key.Replace("`t", " ").Replace("`n", " ").Replace("`r", " "))              # Show name
        $null = $stringBuilder.Append('//')                                                                              # //
        $null = $stringBuilder.Append($episode.c01.Replace("`t", " ").Replace("`n", " ").Replace("`r", " "))             # Episode description
        $null = $stringBuilder.Append("`n")                                                                              # New line
        $null = $stringBuilder.Append($episode.c18)                                                                      # File with full path
        $null = $stringBuilder.Append("`n")                                                                              # New line
        $iSubCount++
        # Progress
        if (-Not($sShows -match ("Season " + $sSeason))) {
            $sShows = $sShows + " | " + ("Season " + $sSeason)
        }   
        if ($sw.Elapsed.TotalMilliseconds -ge 10) {
            $activity3 = ("Generating m3u (" + $iSubCount + "/" + $iTotalCount + ")for:")
            Write-Progress -activity "Totals:"  -status ("Networks: " + $iTotalNetworks + " | Network Episodes: " + $iTotalNetworkEps + " | Shows: " + $iTotalShows + " | Episodes: " + $iTotalShowEps) -Id 1
            Write-Progress -activity $activity3 -status $sShows -Id 3 -percentComplete (($iSubCount / $iTotalCount) * 100)
            $sw.Reset(); $sw.Start()
        }
    }
    $sw = $null

    # Add m3u data to $aM3u
    $sM3uEntry = $stringBuilder.ToString()
    $stringBuilder = $null
    $aM3u.Add('channel_' + $iCount + '.m3u', $sM3uEntry)
    $iCount++
}

# Parse movie genre channels and enter them in to settings2.xml temp string
$iCount = 900
$iCurCh = $iCount
$iChannelPerGenre = 3
foreach ($genreCh in $aGenreCh.GetEnumerator() | Sort-Object Name) {
    if ([string]::IsNullOrEmpty($genreCh.key)) {continue}		                                         # Skip blanks
    for ($i = 1; $i -le $iChannelPerGenre; $i++) {
        $sS2XML += ("    <setting id=`"Channel_" + $iCount + "_type`" value=`"0`" />")
        $sS2XML += "`n"
        $sS2XML += ("    <setting id=`"Channel_" + $iCount + "_1`" value=`"special://profile/playlists/video/" + $genreCh.key + ".xsp`" />")
        $sS2XML += "`n"
        $sS2XML += ("    <setting id=`"Channel_" + $iCount + "_changed`" value=`"False`" />")
        $sS2XML += "`n"
        $sS2XML += ("    <setting id=`"Channel_" + $iCount + "_time`" value=`"1`" />")
        $sS2XML += "`n"

        $activity2 = "Scanning Movie Genres"
        Write-Progress -activity $activity2 -status $genreCh.key -Id 2 -percentComplete ((($iCount - $iCurCh) / ($aGenreCh.Count * $iChannelPerGenre)) * 100)

        $iSubCount = 0
        $iTotalCount = $genreCh.value.Count * $iChannelPerGenre
        $sw = [System.Diagnostics.Stopwatch]::StartNew()
    
        $stringBuilder = New-Object System.Text.StringBuilder
        $sortedgenre = $genreCh.value | Get-Random -Count ([int]::MaxValue)                             # Randomize network episodes if more than 1 show is present on the network
        foreach ($GenreSet in $sortedgenre) {
            if ([string]::IsNullOrEmpty($GenreSet.idMovie)) {continue}
            if (([int]$GenreSet.c11 -lt 100) -or ([int]$GenreSet.c11 -gt 100000)) {
                $MediaLength = Get-MediaLength($GenreSet.c22)
                if ($MediaLength -ne 0) {
                    # Update cache in db at c11
                    $sTempQuery = $sQueryUpdateMovieLength.Replace("???LENGTH???", $MediaLength).Replace("???ID???", $GenreSet.idMovie)
                    Invoke-MySQLNonQuery $conn $sTempQuery
                }
            }
            else {
                # Use cached media length/duration
                $MediaLength = $GenreSet.c11
            }
            $null = $stringBuilder.Append('#EXTINF:')                                                                        # #EXTINF:
            $null = $stringBuilder.Append($MediaLength)                                                                      # Media length/duration
            $null = $stringBuilder.Append(',')                                                                               # ,
            $null = $stringBuilder.Append($GenreSet.c00.Replace("`t", " ").Replace("`n", " ").Replace("`r", " "))            # Movie name
            $null = $stringBuilder.Append('////')                                                                            # ////
            $null = $stringBuilder.Append($GenreSet.c01.Replace("`t", " ").Replace("`n", " ").Replace("`r", " "))            # Movie description
            $null = $stringBuilder.Append("`n")                                                                              # New line
            $null = $stringBuilder.Append($GenreSet.c22)                                                                     # File with full path
            $null = $stringBuilder.Append("`n")                                                                              # New line
            $iSubCount++
            # Progress
            if ($sw.Elapsed.TotalMilliseconds -ge 100) {
                $activity3 = ("Generating m3u (" + $iSubCount + "/" + $iTotalCount + ")for:")
                Write-Progress -activity $activity3 -status $GenreSet.c00 -Id 3 -percentComplete (($iSubCount / $iTotalCount) * 100)
                $sw.Reset(); $sw.Start()
            }
        }
        $sw = $null
        # Add m3u data to $aM3u
        $sM3uEntry = $stringBuilder.ToString()
        $stringBuilder = $null
        $aM3u.Add('channel_' + $iCount + '.m3u', $sM3uEntry)
        $iCount++
    }
}

# Parse movie decade channels and enter them in to settings2.xml temp string
$iCurCh = $iCount
$iChannelPerDecade = 3
foreach ($decadeCh in $aDecadeCh.GetEnumerator() | Sort-Object Name) {
    if ([string]::IsNullOrEmpty($decadeCh.key)) {continue}		                                         # Skip blanks
    for ($i = 1; $i -le $iChannelPerDecade; $i++) {
        $sS2XML += ("    <setting id=`"Channel_" + $iCount + "_type`" value=`"0`" />")
        $sS2XML += "`n"
        $sS2XML += ("    <setting id=`"Channel_" + $iCount + "_1`" value=`"special://profile/playlists/video/" + $decadeCh.key + ".xsp`" />")
        $sS2XML += "`n"
        $sS2XML += ("    <setting id=`"Channel_" + $iCount + "_changed`" value=`"False`" />")
        $sS2XML += "`n"
        $sS2XML += ("    <setting id=`"Channel_" + $iCount + "_time`" value=`"1`" />")
        $sS2XML += "`n"

        $activity2 = "Scanning Movie Decades"
        Write-Progress -activity $activity2 -status $decadeCh.key -Id 2 -percentComplete ((($iCount - $iCurCh) / ($aDecadeCh.Count * $iChannelPerDecade)) * 100)

        $iSubCount = 0
        $iTotalCount = $decadeCh.value.Count * $iChannelPerDecade
        $sw = [System.Diagnostics.Stopwatch]::StartNew()
    
        $stringBuilder = New-Object System.Text.StringBuilder
        $sorteddecade = $decadeCh.value | Get-Random -Count ([int]::MaxValue)                             # Randomize network episodes if more than 1 show is present on the network
        foreach ($DecadeSet in $sorteddecade) {
            if ([string]::IsNullOrEmpty($DecadeSet.idMovie)) {continue}
            if (([int]$DecadeSet.c11 -lt 100) -or ([int]$DecadeSet.c11 -gt 100000)) {
                $MediaLength = Get-MediaLength($DecadeSet.c22)
                if ($MediaLength -ne 0) {
                    # Update cache in db at c11
                    $sTempQuery = $sQueryUpdateMovieLength.Replace("???LENGTH???", $MediaLength).Replace("???ID???", $DecadeSet.idMovie)
                    Invoke-MySQLNonQuery $conn $sTempQuery
                }
            }
            else {
                # Use cached media length/duration
                $MediaLength = $DecadeSet.c11
            }
            $null = $stringBuilder.Append('#EXTINF:')                                                                        # #EXTINF:
            $null = $stringBuilder.Append($MediaLength)                                                                      # Media length/duration
            $null = $stringBuilder.Append(',')                                                                               # ,
            $null = $stringBuilder.Append($DecadeSet.c00.Replace("`t", " ").Replace("`n", " ").Replace("`r", " "))            # Movie name
            $null = $stringBuilder.Append('////')                                                                            # ////
            $null = $stringBuilder.Append($DecadeSet.c01.Replace("`t", " ").Replace("`n", " ").Replace("`r", " "))            # Movie description
            $null = $stringBuilder.Append("`n")                                                                              # New line
            $null = $stringBuilder.Append($DecadeSet.c22)                                                                     # File with full path
            $null = $stringBuilder.Append("`n")                                                                              # New line
            $iSubCount++
            # Progress
            if ($sw.Elapsed.TotalMilliseconds -ge 100) {
                $activity3 = ("Generating m3u (" + $iSubCount + "/" + $iTotalCount + ")for:")
                Write-Progress -activity $activity3 -status $DecadeSet.c00 -Id 3 -percentComplete (($iSubCount / $iTotalCount) * 100)
                $sw.Reset(); $sw.Start()
            }
        }
        $sw = $null
        # Add m3u data to $aM3u
        $sM3uEntry = $stringBuilder.ToString()
        $stringBuilder = $null
        $aM3u.Add('channel_' + $iCount + '.m3u', $sM3uEntry)
        $iCount++
    }
}

# Create genre XSPs
foreach ($genreCh in $aGenreCh.GetEnumerator() | Sort-Object Name) {
    if ([string]::IsNullOrEmpty($genreCh.key)) {continue}		                                         # Skip blanks
    $sXSP = ""
    $sXSP += '<?xml version="1.0" encoding="UTF-8" standalone="yes" ?>'
    $sXSP += "`n"
    $sXSP += '<smartplaylist type="movies">'
    $sXSP += "`n"
    $sXSP += '    <name>' + $genreCh.key + '</name>'
    $sXSP += "`n"
    $sXSP += '    <match>all</match>'
    $sXSP += "`n"
    $sXSP += '    <rule field="path" operator="doesnotcontain">'
    $sXSP += "`n"
    $sXSP += '        <value>/Media/Video/Ad</value>'
    $sXSP += "`n"
    $sXSP += '    </rule>'
    $sXSP += "`n"
    $sXSP += '    <rule field="genre" operator="contains">'
    $sXSP += "`n"
    $sXSP += '        <value>' + $genreCh.key + '</value>'
    $sXSP += "`n"
    $sXSP += '    </rule>'
    $sXSP += "`n"
    $sXSP += '    <group>none</group>'
    $sXSP += "`n"
    $sXSP += '    <order direction="ascending">random</order>'
    $sXSP += "`n"
    $sXSP += '</smartplaylist>'
    $sXSP += "`n"
    $aXSP.Add($genrech.Key + ".xsp", $sXSP)
}

# Create decade XSPs
foreach ($decadeCh in $aDecadeCh.GetEnumerator() | Sort-Object Name) {
    if ([string]::IsNullOrEmpty($decadeCh.key)) {continue}		                                         # Skip blanks
    $sXSP = ""
    $sXSP += '<?xml version="1.0" encoding="UTF-8" standalone="yes" ?>'
    $sXSP += "`n"
    $sXSP += '<smartplaylist type="movies">'
    $sXSP += "`n"
    $sXSP += '    <name>' + ($decadeCh.key) + 's' + '</name>'
    $sXSP += "`n"
    $sXSP += '    <match>all</match>'
    $sXSP += "`n"
    $sXSP += '    <rule field="path" operator="doesnotcontain">'
    $sXSP += "`n"
    $sXSP += '        <value>/Media/Video/Ad</value>'
    $sXSP += "`n"
    $sXSP += '    </rule>'
    $sXSP += "`n"
    $sXSP += '    <rule field="year" operator="greaterthan">'
    $sXSP += "`n"
    $sXSP += '        <value>' + ([int]$decadeCh.key - 1) + '</value>'
    $sXSP += "`n"
    $sXSP += '    </rule>'
    $sXSP += "`n"
    $sXSP += '    <rule field="year" operator="lessthan">'
    $sXSP += "`n"
    $sXSP += '        <value>' + ([int]$decadeCh.key + 10) + '</value>'
    $sXSP += "`n"
    $sXSP += '    </rule>'
    $sXSP += "`n"
    $sXSP += '    <order direction="ascending">random</order>'
    $sXSP += "`n"
    $sXSP += '</smartplaylist>'
    $sXSP += "`n"
    $aXSP.Add($decadeCh.key + ".xsp", $sXSP)
}




# Add additional XSPs
#Ad filtered
$sXSP = ""
$sXSP += '<?xml version="1.0" encoding="UTF-8" standalone="yes" ?>'
$sXSP += "`n"
$sXSP += '<smartplaylist type="movies">'
$sXSP += "`n"
$sXSP += '    <name>Movies</name>'
$sXSP += "`n"
$sXSP += '    <match>all</match>'
$sXSP += "`n"
$sXSP += '    <rule field="path" operator="doesnotcontain">'
$sXSP += "`n"
$sXSP += '        <value>/Media/Video/Ad</value>'
$sXSP += "`n"
$sXSP += '    </rule>'
$sXSP += "`n"
$sXSP += '    <group>none</group>'
$sXSP += "`n"
$sXSP += '    <order direction="ascending">sorttitle</order>'
$sXSP += "`n"
$sXSP += '</smartplaylist>'
$sXSP += "`n"
$aXSP.Add("Movies.xsp", $sXSP)
#Ad
$sXSP = ""
$sXSP += '<?xml version="1.0" encoding="UTF-8" standalone="yes" ?>'
$sXSP += "`n"
$sXSP += '<smartplaylist type="movies">'
$sXSP += "`n"
$sXSP += '    <name>Adult</name>'
$sXSP += "`n"
$sXSP += '    <match>all</match>'
$sXSP += "`n"
$sXSP += '    <rule field="path" operator="contains">'
$sXSP += "`n"
$sXSP += '        <value>/Media/Video/Ad</value>'
$sXSP += "`n"
$sXSP += '    </rule>'
$sXSP += "`n"
$sXSP += '    <group>none</group>'
$sXSP += "`n"
$sXSP += '    <order direction="ascending">sorttitle</order>'
$sXSP += "`n"
$sXSP += '</smartplaylist>'
$sXSP += "`n"
$aXSP.Add("Adult.xsp", $sXSP)
#All movies
$sXSP = ""
$sXSP += '<?xml version="1.0" encoding="UTF-8" standalone="yes" ?>'
$sXSP += "`n"
$sXSP += '<smartplaylist type="movies">'
$sXSP += "`n"
$sXSP += '    <name>All Movies</name>'
$sXSP += "`n"
$sXSP += '    <match>any</match>'
$sXSP += "`n"
$sXSP += '    <group>none</group>'
$sXSP += "`n"
$sXSP += '    <order direction="ascending">sorttitle</order>'
$sXSP += "`n"
$sXSP += '</smartplaylist>'
$sXSP += "`n"
$aXSP.Add("All Movies.xsp", $sXSP)

# Closing xml string
$sS2XML += "    <setting id=`"LastResetTime`" value=`"1495239376`" />"
$sS2XML += "`n"
$sS2XML += "    <setting id=`"LastExitTime`" value=`"1495239391`" />"
$sS2XML += "`n"
$sS2XML += "</settings>"

# Write settings.xml
$sSXML = ''
$sSXML += '<settings>'
$sSXML += '    <setting id="AutoOff" value="0" />'
$sSXML += '    <setting id="ChannelDelay" value="1" />'
$sSXML += '    <setting id="ChannelLogoFolder" value="special://home/addons/script.pseudotv/resources/logos/" />'
$sSXML += '    <setting id="ChannelResetSetting" value="4" />'
$sSXML += '    <setting id="ChannelSharing" value="false" />'
$sSXML += '    <setting id="ClipLength" value="4" />'
$sSXML += '    <setting id="ClockMode" value="0" />'
$sSXML += '    <setting id="ConfigDialog" value="" />'
$sSXML += '    <setting id="CurrentChannel" value="1" />'
$sSXML += '    <setting id="EnableComingUp" value="false" />'
$sSXML += '    <setting id="ForceChannelReset" value="false" />'
$sSXML += '    <setting id="HideClips" value="true" />'
$sSXML += '    <setting id="InfoOnChange" value="false" />'
$sSXML += '    <setting id="MediaLimit" value="7" />'
$sSXML += '    <setting id="NumberColour" value="2" />'
$sSXML += '    <setting id="SeekBackward" value="1" />'
$sSXML += '    <setting id="SeekForward" value="1" />'
$sSXML += '    <setting id="SettingsFolder" value="" />'
$sSXML += '    <setting id="ShowChannelBug" value="false" />'
$sSXML += '    <setting id="ShowEpgLogo" value="true" />'
$sSXML += '    <setting id="ShowSeEp" value="true" />'
$sSXML += '    <setting id="StartMode" value="2" />'
$sSXML += '    <setting id="ThreadMode" value="1" />'
$sSXML += '    <setting id="enable" value="false" />'
$sSXML += '    <setting id="notify" value="true" />'
$sSXML += '    <setting id="timer_amount" value="1" />'
$sSXML += '</settings>'

# Write settings.xml to temp dir
$sFile = ($env:TEMP + "\PTV\settings.xml")
$Utf8NoBomEncoding = New-Object System.Text.UTF8Encoding $False
[System.IO.File]::WriteAllLines($sFile, $sSXML, $Utf8NoBomEncoding)

# Write settings2.xml to temp dir
$sFile = ($env:TEMP + "\PTV\settings2.xml")
$Utf8NoBomEncoding = New-Object System.Text.UTF8Encoding $False
[System.IO.File]::WriteAllLines($sFile, $sS2XML, $Utf8NoBomEncoding)

# Write M3U Files to disk
foreach ($file in $aM3u.GetEnumerator()) {
    $sFile = ($env:TEMP + "\PTV\cache\" + $file.key)
    $sTemp = "#EXTM3U"
    $sTemp += "`n"
    $sTemp += $file.Value

    $Utf8NoBomEncoding = New-Object System.Text.UTF8Encoding $False
    [System.IO.File]::WriteAllLines($sFile, $sTemp, $Utf8NoBomEncoding)

    #$sTemp | Out-File -FilePath $sFile -Force
}

# Write XSP Files to disk
foreach ($file in $aXSP.GetEnumerator()) {
    $sFile = ($env:TEMP + "\MoviePlaylists\" + $file.key)

    $Utf8NoBomEncoding = New-Object System.Text.UTF8Encoding $False
    [System.IO.File]::WriteAllLines($sFile, $file.Value, $Utf8NoBomEncoding)

    #$file.Value | Out-File -FilePath $sFile -Force
}

# Temp file management (remove old temp dirs and recreate blank structure)
$aPaths = @()                                                                                           # Array of paths to remove and recreate empty
$aPaths += , ($env:APPDATA + "\Kodi\userdata\playlists\video")
$aPaths += , ($env:APPDATA + "\Kodi\userdata\addon_data\script.pseudotv")
foreach ($sPath in $aPaths) {
    if (Test-Path -PathType Container $sPath) {
        # $sPath already exists
        Remove-Item $sPath -Force -Recurse                                                              # Delete $sPath
        New-Item -ItemType Directory -Force -Path $sPath | Out-Null                                     # Recreate $sPath
    }
    else {
        # $sPath doesn't exist
        New-Item -ItemType Directory -Force -Path $sPath | Out-Null                                     # Create $sPath
    }
}

# Copy files
#Copy-Item ($env:TEMP + "\PTV\settings2.xml") -Destination ($env:APPDATA + "\Kodi\userdata\addon_data\script.pseudotv")
Copy-Item ($env:TEMP + "\PTV\*") -Destination ($env:APPDATA + "\Kodi\userdata\addon_data\script.pseudotv") -Recurse
Copy-Item ($env:TEMP + "\MoviePlaylists\*.xsp") -Destination ($env:APPDATA + "\Kodi\userdata\playlists\video") -Recurse

# Cleanup
Disconnect-MySQL($conn)

#Wait
Write-Progress -Activity $activity2 -Status "Ready" -Id 2 -Completed
Write-Progress -Activity $activity3 -Status "Ready" -Id 3 -Completed

$sWeaselTV = "`n`n`n`n`n`n`n`n                                                                             ,----,           
                                                                           ,/   .``|           
             .---.                                               ,--,    ,``   .'  :           
            /. ./|                                             ,--.'|  ;    ;     /     ,---. 
        .--'.  ' ;                                             |  | :.'___,/    ,'     /__./| 
       /__./ \ : |                         .--.--.             :  : '|    :     | ,---.;  ; | 
   .--'.  '   \' .   ,---.     ,--.--.    /  /    '     ,---.  |  ' |;    |.';  ;/___/ \  | | 
  /___/ \ |    ' '  /     \   /       \  |  :  /``./    /     \ '  | |``----'  |  |\   ;  \ ' | 
  ;   \  \;      : /    /  | .--.  .-. | |  :  ;_     /    /  ||  | :    '   :  ; \   \  \: | 
   \   ;  ``      |.    ' / |  \__\/: . .  \  \    ``. .    ' / |'  : |__  |   |  '  ;   \  ' . 
    .   \    .\  ;'   ;   /|  ,`" .--.; |   ``----.   \'   ;   /||  | '.'| '   :  |   \   \   ' 
     \   \   ' \ |'   |  / | /  /  ,.  |  /  /``--'  /'   |  / |;  :    ; ;   |.'     \   ``  ; 
      :   '  |--`" |   :    |;  :   .'   \'--'.     / |   :    ||  ,   /  '---'        :   \ | 
       \   \ ;     \   \  / |  ,     .-./  ``--'---'   \   \  /  ---``-'                 '---`"  
        '---`"       ``----'   ``--``---'                  ``----'                                 
                                                                                              
"
$sWeaselTV

Write-Host -NoNewLine 'Press any key to continue...';
$null = $Host.UI.RawUI.ReadKey('NoEcho,IncludeKeyDown');