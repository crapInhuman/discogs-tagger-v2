Option Explicit
'
' Discogs Batch Tagger Script for MediaMonkey ( crap_inhuman with a little help from my friends Let & eepman )
'
Const VersionStr = "v2.12"

'Changes from 2.11 to 2.12 by crap_inhuman in 03.2014
'		Moving the tracks with the Up and Down Button now work
'		Bug remoeved: Sub-Track do not select(set) the song
'		Added the option for switching the last artist separator ("&" or "chosen separator")

'Changes from 2.10 to 2.11 by crap_inhuman in 03.2014
'		Removed bug with more than one artist for a title
'		Added simple routine to check for false position separators

'Changes from 1.01 to 2.10 by crap_inhuman in 02.2014
'		Keywords are now not case sensitive
'		Added Set Locale for supporting more countries
'		Added the Featuring Keywords
'		Fixed a bug with the new submission form of discogs
'		The script now shows the filtered total and the matched total
'		Raise the max count of release results to 100
'		Display the number of matched releases and which one you are viewing in the search bar
'		and much more..

'Changes from 1.01 to 2.10 by crap_inhuman in 02.2014 (not released)
		'Changed the image access method

'Changes from 1.00 to 1.01 by crap_inhuman in 10.2013 (not released)
'		Removed bug in extra artist assignment
'		Added 'Don't save' and 4 more fields for saving release-number
'		Added 2 options: Process only Discogs releases, Process no Discogs releases
'		Options now show at start of script

'First Version 1.00 by crap_inhuman in 09.2013 (not released)
'		The date and original date tags now always been updated, if the option set (e.g. if the date tag at discogs is blank , the date of the tagged album will be blank too
'		Show info messagebox before and after script-usage
'		Removed bug: Sub track name will not recognized if it is the last track



' ToDo:	Add more tooltips to the html
'		Trackauswahl im unteren Track-Fenster aktivieren (Zeile 5009)
'		Falsche Anzahl releasaes prüfen


' WebBrowser is visible browser object with display of discogs album info
Dim WebBrowser, WebBrowser2

' decoded json object representing currently selected release
Dim CurrentRelease, ReleaseSkip

Dim Form

Dim cTab : cTab = 1

Dim UI, ini

Dim ResultsReleaseID ' result list
Dim CurrentReleaseID
Dim tracklistHTML

Dim templateHTML
Dim Combo, Head, btn2, SearchFormWidth

Dim Tracks, ArtistTitles, TracksNum, TracksCD, InvolvedArtists, Lyricists, Composers, Conductors, Producers
Dim AlbumTitle, AlbumArtist, AlbumArtistTitle, AlbumIDList
Dim theCatalogs, theCountry, theFormat, theLabels, Styles, Genres, Comment

Dim CheckAlbum, CheckArtist, CheckAlbumArtist, CheckAlbumArtistFirst, CheckLabel, CheckDate, CheckOrigDate, CheckGenre
Dim CheckCountry, CheckCover, CheckSmallCover, CheckStyle, CheckCatalog, CheckRelease, CheckInvolved, CheckLyricist
Dim CheckComposer, CheckConductor, CheckProducer, CheckDiscNum, CheckTrackNum, CheckFormat, CheckUseAnv, CheckYearOnlyDate
Dim CheckForceNumeric, CheckSidesToDisc, CheckForceDisc, CheckNoDisc, CheckLeadingZero, CheckVarious, TxtVarious
Dim CheckTitleFeaturing, CheckComment, CheckFeaturingName, TxtFeaturingName, CheckOriginalDiscogsTrack
Dim CheckUnselectNoTrackPos, CheckStyleField, CheckNotAlwaysSaveImage
Dim SubTrackNameSelection
Dim CountryFilterList, MediaTypeFilterList, MediaFormatFilterList, YearFilterList
Dim LyricistKeywords, ConductorKeywords, ProducerKeywords, ComposerKeywords, FeaturingKeywords
Dim ArtistSeparator, ArtistLastSeparator
Dim RadioBoxCheck

Dim SavedReleaseId
Dim SavedSearchTerm
Dim SavedMasterId, SavedArtistId, SavedLabelId

Dim FilterMediaType, FilterCountry, FilterYear, FilterMediaFormat, CurrentLoadType
Dim MediaTypeList, MediaFormatList, CountryList, YearList, AlternativeList, LoadList

Dim FirstTrack
Dim AlbumArtURL, AlbumArtThumbNail
Dim iMaxTracks
Dim iAutoTrackNumber, iAutoDiscNumber
Dim LastDisc
Dim SelectAll, UnselectedTracks(1000)

Dim ReleaseTag, CountryTag, CatalogTag, FormatTag
Dim OriginalDate, ReleaseDate, Separator
Dim UserChoose

Dim CurrentSelectedAlbum
Dim fso, loc, logf
Dim SkipNotChangedReleases, ProcessOnlyDiscogs, ProcessNoDiscogs

Dim cReleasesUpdate, cReleasesSkip, cReleasesAutoSkip, cReleasesOnlyDiscogsSkip, cReleasesNoDiscogsSkip
Dim NewTrackList, SongList, WebTrackCount
Dim ErrorMessage


'----------------------------------DiscogsImages----------------------------------------
Rem Dim SaveImageType, SaveImage, CoverStorage, FileNameList
Rem Dim ImageTypeList, ImageList
Rem Dim list
Rem Dim ImagesCount
Rem Dim SaveMoreImages
Rem Dim WebBrowser3
Rem Dim SelectedSongsGlobal
'----------------------------------DiscogsImages----------------------------------------

Rem Dim SDB : Set SDB = CreateObject("SongsDB.SDBApplication")
Rem SDB.ShutdownAfterDisconnect = False
Rem Dim Script : Set Script = CreateObject("SongsDB.SDBScriptControl")
Set UI = SDB.UI
Rem Call BatchDiscogsSearch

' MediaMonkey calls this method whenever a search is started using this script
Sub BatchDiscogsSearch()
	
	WriteLogInit
	
	WriteLog "Start BatchDiscogsSearch"
	
	cReleasesUpdate = 0
	cReleasesSkip = 0
	cReleasesAutoSkip = 0
	cReleasesOnlyDiscogsSkip = 0
	cReleasesNoDiscogsSkip = 0
	ReleaseSkip = False
	
	Dim tmpCountry, tmpCountry2, tmpMediaType, tmpMediaType2, tmpMediaFormat, tmpMediaFormat2, tmpYear, tmpYear2
	Dim i, a, tmp
	Set CountryFilterList = SDB.NewStringList
	Set MediaTypeFilterList = SDB.NewStringList
	Set MediaFormatFilterList = SDB.NewStringList
	Set YearFilterList = SDB.NewStringList
	
	'*FilterList.Item(0) = "0" -> No Filter
	'*FilterList.Item(0) = "1" -> Custom Filter
	'*FilterList.Item(0) = "2" -> Selected Country/MediaType/MediaFormat/Year
	
	Set ini = SDB.IniFile
	If Not (ini Is Nothing) Then
		If ini.StringValue("DiscogsAutoTagWeb","CheckAlbum") = "" Then
			ini.BoolValue("DiscogsAutoTagWeb","CheckAlbum") = True
		End If
		If ini.StringValue("DiscogsAutoTagWeb","CheckArtist") = "" Then
			ini.BoolValue("DiscogsAutoTagWeb","CheckArtist") = True
		End If
		If ini.StringValue("DiscogsAutoTagWeb","CheckAlbumArtist") = "" Then
			ini.BoolValue("DiscogsAutoTagWeb","CheckAlbumArtist") = True
		End If
		If ini.StringValue("DiscogsAutoTagWeb","CheckAlbumArtistFirst") = "" Then
			ini.BoolValue("DiscogsAutoTagWeb","CheckAlbumArtistFirst") = True
		End If
		If ini.StringValue("DiscogsAutoTagWeb","CheckLabel") = "" Then
			ini.BoolValue("DiscogsAutoTagWeb","CheckLabel") = True
		End If
		If ini.StringValue("DiscogsAutoTagWeb","CheckDate") = "" Then
			ini.BoolValue("DiscogsAutoTagWeb","CheckDate") = True
		End If
		If ini.StringValue("DiscogsAutoTagWeb","CheckOrigDate") = "" Then
			ini.BoolValue("DiscogsAutoTagWeb","CheckOrigDate") = False
		End If
		If ini.StringValue("DiscogsAutoTagWeb","CheckGenre") = "" Then
			ini.BoolValue("DiscogsAutoTagWeb","CheckGenre") = False
		End If
		If ini.StringValue("DiscogsAutoTagWeb","CheckStyle") = "" Then
			ini.BoolValue("DiscogsAutoTagWeb","CheckStyle") = True
		End If
		If ini.StringValue("DiscogsAutoTagWeb","CheckCountry") = "" Then
			ini.BoolValue("DiscogsAutoTagWeb","CheckCountry") = True
		End If
		If ini.StringValue("DiscogsAutoTagWeb","CheckCover") = "" Then
			ini.BoolValue("DiscogsAutoTagWeb","CheckCover") = True
		End If
		If ini.StringValue("DiscogsAutoTagWeb","CheckSmallCover") = "" Then
			ini.BoolValue("DiscogsAutoTagWeb","CheckSmallCover") = False
		End If
		If ini.StringValue("DiscogsAutoTagWeb","CheckCatalog") = "" Then
			ini.BoolValue("DiscogsAutoTagWeb","CheckCatalog") = True
		End If
		If ini.StringValue("DiscogsAutoTagWeb","CheckRelease") = "" Then
			ini.BoolValue("DiscogsAutoTagWeb","CheckRelease") = True
		End If
		If ini.StringValue("DiscogsAutoTagWeb","CheckInvolved") = "" Then
			ini.BoolValue("DiscogsAutoTagWeb","CheckInvolved") = False
		End If
		If ini.StringValue("DiscogsAutoTagWeb","CheckLyricist") = "" Then
			ini.BoolValue("DiscogsAutoTagWeb","CheckLyricist") = False
		End If
		If ini.StringValue("DiscogsAutoTagWeb","CheckComposer") = "" Then
			ini.BoolValue("DiscogsAutoTagWeb","CheckComposer") = False
		End If
		If ini.StringValue("DiscogsAutoTagWeb","CheckConductor") = "" Then
			ini.BoolValue("DiscogsAutoTagWeb","CheckConductor") = False
		End If
		If ini.StringValue("DiscogsAutoTagWeb","CheckProducer") = "" Then
			ini.BoolValue("DiscogsAutoTagWeb","CheckProducer") = False
		End If
		If ini.StringValue("DiscogsAutoTagWeb","CheckDiscNum") = "" Then
			ini.BoolValue("DiscogsAutoTagWeb","CheckDiscNum") = True
		End If
		If ini.StringValue("DiscogsAutoTagWeb","CheckTrackNum") = "" Then
			ini.BoolValue("DiscogsAutoTagWeb","CheckTrackNum") = True
		End If
		If ini.StringValue("DiscogsAutoTagWeb","CheckFormat") = "" Then
			ini.BoolValue("DiscogsAutoTagWeb","CheckFormat") = True
		End If
		If ini.StringValue("DiscogsAutoTagWeb","CheckUseAnv") = "" Then
			ini.BoolValue("DiscogsAutoTagWeb","CheckUseAnv") = False
		End If
		If ini.StringValue("DiscogsAutoTagWeb","CheckYearOnlyDate") = "" Then
			ini.BoolValue("DiscogsAutoTagWeb","CheckYearOnlyDate") = False
		End If
		If ini.StringValue("DiscogsAutoTagWeb","CheckForceNumeric") = "" Then
			ini.BoolValue("DiscogsAutoTagWeb","CheckForceNumeric") = False
		End If
		If ini.StringValue("DiscogsAutoTagWeb","CheckSidesToDisc") = "" Then
			ini.BoolValue("DiscogsAutoTagWeb","CheckSidesToDisc") = False
		End If
		If ini.StringValue("DiscogsAutoTagWeb","CheckForceDisc") = "" Then
			ini.BoolValue("DiscogsAutoTagWeb","CheckForceDisc") = False
		End If
		If ini.StringValue("DiscogsAutoTagWeb","CheckNoDisc") = "" Then
			ini.BoolValue("DiscogsAutoTagWeb","CheckNoDisc") = False
		End If
		If ini.StringValue("DiscogsAutoTagWeb","CheckOriginalDiscogsTrack") = "" Then
			ini.BoolValue("DiscogsAutoTagWeb","CheckOriginalDiscogsTrack") = False
		End If
		If ini.StringValue("DiscogsAutoTagWeb","ReleaseTag") = "" Then
			ini.StringValue("DiscogsAutoTagWeb","ReleaseTag") = "Custom2"
		End If
		If ini.StringValue("DiscogsAutoTagWeb","CatalogTag") = "" Then
			ini.StringValue("DiscogsAutoTagWeb","CatalogTag") = "Custom3"
		End If
		If ini.StringValue("DiscogsAutoTagWeb","CountryTag") = "" Then
			ini.StringValue("DiscogsAutoTagWeb","CountryTag") = "Custom4"
		End If
		If ini.StringValue("DiscogsAutoTagWeb","FormatTag") = "" Then
			ini.StringValue("DiscogsAutoTagWeb","FormatTag") = "Custom5"
		End If
		If ini.StringValue("DiscogsAutoTagWeb","CheckLeadingZero") = "" Then
			ini.BoolValue("DiscogsAutoTagWeb","CheckLeadingZero") = True
		End If
		If ini.StringValue("DiscogsAutoTagWeb","CheckVarious") = "" Then
			ini.BoolValue("DiscogsAutoTagWeb","CheckVarious") = False
		End If
		If ini.StringValue("DiscogsAutoTagWeb","TxtVarious") = "" Then
			ini.StringValue("DiscogsAutoTagWeb","TxtVarious") = "Various Artists"
		End If
		If ini.StringValue("DiscogsAutoTagWeb","CheckTitleFeaturing") = "" Then
			ini.BoolValue("DiscogsAutoTagWeb","CheckTitleFeaturing") = True
		End If
		If ini.StringValue("DiscogsAutoTagWeb","CheckFeaturingName") = "" Then
			ini.BoolValue("DiscogsAutoTagWeb","CheckFeaturingName") = True
		End If
		If ini.StringValue("DiscogsAutoTagWeb","TxtFeaturingName") = "" Then
			ini.StringValue("DiscogsAutoTagWeb","TxtFeaturingName") = "feat."
		End If
		If ini.StringValue("DiscogsAutoTagWeb","CheckComment") = "" Then
			ini.BoolValue("DiscogsAutoTagWeb","CheckComment") = True
		End If
		If ini.StringValue("DiscogsAutoTagWeb","CheckUnselectNoTrackPos") = "" Then
			ini.BoolValue("DiscogsAutoTagWeb","CheckUnselectNoTrackPos") = True
		End If
		If ini.StringValue("DiscogsAutoTagWeb","SubTrackNameSelection") = "" Then
			ini.BoolValue("DiscogsAutoTagWeb","SubTrackNameSelection") = False
		End If
		If ini.StringValue("DiscogsAutoTagWeb","CurrentCountryFilter") = "" Then
			tmp = "0"
			For a = 1 To 282
				tmp = tmp & ",0"
			Next
			ini.StringValue("DiscogsAutoTagWeb","CurrentCountryFilter") = tmp
		End If
		
		If ini.StringValue("DiscogsAutoTagWeb","CurrentMediaTypeFilter") = "" Then
			tmp = "0"
			For a = 1 To 38
				tmp = tmp & ",0"
			Next
			ini.StringValue("DiscogsAutoTagWeb","CurrentMediaTypeFilter") = tmp
		End If
		
		If ini.StringValue("DiscogsAutoTagWeb","CurrentMediaFormatFilter") = "" Then
			tmp = "0"
			For a = 1 To 48
				tmp = tmp & ",0"
			Next
			ini.StringValue("DiscogsAutoTagWeb","CurrentMediaFormatFilter") = tmp
		End If
		
		If ini.StringValue("DiscogsAutoTagWeb","CurrentYearFilter") = "" Then
			tmp = "0"
			For a = Year(Date) To 1900 Step -1
				tmp = tmp & ",0"
			Next
			ini.StringValue("DiscogsAutoTagWeb","CurrentYearFilter") = tmp
		End If
		
		If ini.StringValue("DiscogsAutoTagWeb","LyricistKeywords") = "" Then
			ini.StringValue("DiscogsAutoTagWeb","LyricistKeywords") = "Lyrics By,Words By"
		End If
		If ini.StringValue("DiscogsAutoTagWeb","ConductorKeywords") = "" Then
			ini.StringValue("DiscogsAutoTagWeb","ConductorKeywords") = "Conductor"
		End If
		If ini.StringValue("DiscogsAutoTagWeb","ProducerKeywords") = "" Then
			ini.StringValue("DiscogsAutoTagWeb","ProducerKeywords") = "Producer,Arranged By,Recorded By"
		End If
		If ini.StringValue("DiscogsAutoTagWeb","ComposerKeywords") = "" Then
			ini.StringValue("DiscogsAutoTagWeb","ComposerKeywords") = "Composed By,Score,Written-By,Written By,Music By,Programmed By,Songwriter"
		End If
		If ini.StringValue("DiscogsAutoTagWeb","FeaturingKeywords") = "" Then
			ini.StringValue("DiscogsAutoTagWeb","FeaturingKeywords") = "featuring,feat.,ft.,ft ,feat ,Rap,Rap [Featuring],Vocals [Featuring]"
		End If

		If ini.StringValue("DiscogsAutoTagWeb", "CheckNotAlwaysSaveImage") = "" Then
			ini.BoolValue("DiscogsAutoTagWeb", "CheckNotAlwaysSaveImage") = False
		End If
		If ini.StringValue("DiscogsAutoTagWeb","CheckStyleField") = "" Then
			ini.StringValue("DiscogsAutoTagWeb","CheckStyleField") = "Default (stored with Genre)"
		End If

		If ini.StringValue("DiscogsAutoTagWeb","ArtistSeparator") = "" Then
			ini.StringValue("DiscogsAutoTagWeb","ArtistSeparator") = ", "
		End If

		If ini.StringValue("DiscogsAutoTagWeb","ArtistLastSeparator") = "" Then
			ini.BoolValue("DiscogsAutoTagWeb","ArtistLastSeparator") = True
		End If

		If ini.StringValue("DiscogsAutoTagWeb","SkipNotChangedReleases") = "" Then
			ini.BoolValue("DiscogsAutoTagWeb","SkipNotChangedReleases") = True
		End If
		
		If ini.StringValue("DiscogsAutoTagWeb","ProcessOnlyDiscogs") = "" Then
			ini.BoolValue("DiscogsAutoTagWeb","ProcessOnlyDiscogs") = False
		End If
		
		If ini.StringValue("DiscogsAutoTagWeb","ProcessNoDiscogs") = "" Then
			ini.BoolValue("DiscogsAutoTagWeb","ProcessNoDiscogs") = False
		End If
		
		'----------------------------------DiscogsImages----------------------------------------
		Rem CoverStorage = ini.StringValue("PreviewSettings","DefaultCoverStorage")
		Rem 'Coverstorage = 0 -> Save image to tag (if possible) otherwise save to file folder
		Rem 'Coverstorage = 1 -> Save image to file folder
		Rem 'Coverstorage = 2 -> Save image to cover folder (is deprecated and will not be supported !!)
		Rem 'Coverstorage = 3 -> Save image to tag (if possible) and to file folder
		Rem If CoverStorage = 2 Then
		Rem Call SDB.MessageBox("Discogs Images: Your Cover Storage is not supported by DiscogsImages !",mtError,Array(mbOk))
		Rem Exit Sub
		Rem End If
		'----------------------------------DiscogsImages----------------------------------------
		
	End If
	
	
	
	CheckAlbum = ini.BoolValue("DiscogsAutoTagWeb","CheckAlbum")
	CheckArtist = ini.BoolValue("DiscogsAutoTagWeb","CheckArtist")
	CheckAlbumArtist = ini.BoolValue("DiscogsAutoTagWeb","CheckAlbumArtist")
	CheckAlbumArtistFirst = ini.BoolValue("DiscogsAutoTagWeb","CheckAlbumArtistFirst")
	CheckLabel = ini.BoolValue("DiscogsAutoTagWeb","CheckLabel")
	CheckDate = ini.BoolValue("DiscogsAutoTagWeb","CheckDate")
	CheckOrigDate = ini.BoolValue("DiscogsAutoTagWeb","CheckOrigDate")
	CheckGenre = ini.BoolValue("DiscogsAutoTagWeb","CheckGenre")
	CheckStyle = ini.BoolValue("DiscogsAutoTagWeb","CheckStyle")
	CheckCountry = ini.BoolValue("DiscogsAutoTagWeb","CheckCountry")
	CheckCover = ini.BoolValue("DiscogsAutoTagWeb","CheckCover")
	CheckSmallCover = ini.BoolValue("DiscogsAutoTagWeb","CheckSmallCover")
	CheckCatalog = ini.BoolValue("DiscogsAutoTagWeb","CheckCatalog")
	CheckRelease = ini.BoolValue("DiscogsAutoTagWeb","CheckRelease")
	CheckInvolved = ini.BoolValue("DiscogsAutoTagWeb","CheckInvolved")
	CheckLyricist = ini.BoolValue("DiscogsAutoTagWeb","CheckLyricist")
	CheckComposer = ini.BoolValue("DiscogsAutoTagWeb","CheckComposer")
	CheckConductor = ini.BoolValue("DiscogsAutoTagWeb","CheckConductor")
	CheckProducer = ini.BoolValue("DiscogsAutoTagWeb","CheckProducer")
	CheckDiscNum = ini.BoolValue("DiscogsAutoTagWeb","CheckDiscNum")
	CheckTrackNum = ini.BoolValue("DiscogsAutoTagWeb","CheckTrackNum")
	CheckFormat = ini.BoolValue("DiscogsAutoTagWeb","CheckFormat")
	CheckUseAnv = ini.BoolValue("DiscogsAutoTagWeb","CheckUseAnv")
	CheckYearOnlyDate = ini.BoolValue("DiscogsAutoTagWeb","CheckYearOnlyDate")
	CheckForceNumeric = ini.BoolValue("DiscogsAutoTagWeb","CheckForceNumeric")
	CheckSidesToDisc = ini.BoolValue("DiscogsAutoTagWeb","CheckSidesToDisc")
	CheckForceDisc = ini.BoolValue("DiscogsAutoTagWeb","CheckForceDisc")
	CheckOriginalDiscogsTrack = ini.BoolValue("DiscogsAutoTagWeb","CheckOriginalDiscogsTrack")
	CheckNoDisc = ini.BoolValue("DiscogsAutoTagWeb","CheckNoDisc")
	ReleaseTag = ini.StringValue("DiscogsAutoTagWeb","ReleaseTag")
	CatalogTag = ini.StringValue("DiscogsAutoTagWeb","CatalogTag")
	CountryTag = ini.StringValue("DiscogsAutoTagWeb","CountryTag")
	FormatTag = ini.StringValue("DiscogsAutoTagWeb","FormatTag")
	CheckLeadingZero = ini.BoolValue("DiscogsAutoTagWeb","CheckLeadingZero")
	CheckVarious = ini.BoolValue("DiscogsAutoTagWeb","CheckVarious")
	TxtVarious = ini.StringValue("DiscogsAutoTagWeb","TxtVarious")
	CheckTitleFeaturing = ini.BoolValue("DiscogsAutoTagWeb","CheckTitleFeaturing")
	CheckFeaturingName = ini.boolValue("DiscogsAutoTagWeb","CheckFeaturingName")
	TxtFeaturingName = ini.StringValue("DiscogsAutoTagWeb","TxtFeaturingName")
	CheckComment = ini.BoolValue("DiscogsAutoTagWeb","CheckComment")
	CheckUnselectNoTrackPos = ini.BoolValue("DiscogsAutoTagWeb","CheckUnselectNoTrackPos")
	SubTrackNameSelection = ini.BoolValue("DiscogsAutoTagWeb","SubTrackNameSelection")
	Separator = ini.StringValue("Appearance","MultiStringSeparator")
	tmpCountry = ini.StringValue("DiscogsAutoTagWeb","CurrentCountryFilter")
	tmpCountry2 = Split(tmpCountry, ",")
	tmpMediaType = ini.StringValue("DiscogsAutoTagWeb","CurrentMediaTypeFilter")
	tmpMediaType2 = Split(tmpMediaType, ",")
	tmpMediaFormat = ini.StringValue("DiscogsAutoTagWeb","CurrentMediaFormatFilter")
	tmpMediaFormat2 = Split(tmpMediaFormat, ",")
	tmpYear = ini.StringValue("DiscogsAutoTagWeb","CurrentYearFilter")
	tmpYear2 = Split(tmpYear, ",")
	LyricistKeywords = ini.StringValue("DiscogsAutoTagWeb","LyricistKeywords")
	ConductorKeywords = ini.StringValue("DiscogsAutoTagWeb","ConductorKeywords")
	ProducerKeywords = ini.StringValue("DiscogsAutoTagWeb","ProducerKeywords")
	ComposerKeywords = ini.StringValue("DiscogsAutoTagWeb","ComposerKeywords")
	FeaturingKeywords = ini.StringValue("DiscogsAutoTagWeb","FeaturingKeywords")
	CheckNotAlwaysSaveImage = ini.BoolValue("DiscogsAutoTagWeb","CheckNotAlwaysSaveImage")
	CheckStyleField = ini.StringValue("DiscogsAutoTagWeb","CheckStyleField")
	ArtistSeparator = ini.StringValue("DiscogsAutoTagWeb","ArtistSeparator")
	ArtistLastSeparator = ini.BoolValue("DiscogsAutoTagWeb","ArtistLastSeparator")
	SkipNotChangedReleases = ini.BoolValue("DiscogsAutoTagWeb","SkipNotChangedReleases")
	ProcessNoDiscogs = ini.BoolValue("DiscogsAutoTagWeb","ProcessNoDiscogs")
	ProcessOnlyDiscogs = ini.BoolValue("DiscogsAutoTagWeb","ProcessOnlyDiscogs")
	
	Separator = Left(Separator, Len(Separator)-1)
	Separator = Right(Separator, Len(Separator)-1)
	
	SelectAll = True
	
	Set MediaTypeList = SDB.NewStringList
	Set MediaFormatList = SDB.NewStringList
	Set CountryList = SDB.NewStringList
	Set YearList = SDB.NewStringList
	Set AlternativeList = SDB.NewStringList
	Set LoadList = SDB.NewStringList
	
	LoadList.Add "Search Results"
	LoadList.Add "Master Release"
	LoadList.Add "Releases of Artist"
	LoadList.Add "Releases of Label"
	
	MediaTypeList.Add "None"
	MediaTypeList.Add "Vinyl"
	MediaTypeList.Add "CD"
	MediaTypeList.Add "DVD"
	MediaTypeList.Add "Blu-Ray"
	MediaTypeList.Add "Cassette"
	MediaTypeList.Add "DAT"
	MediaTypeList.Add "Minidisc"
	MediaTypeList.Add "File"
	MediaTypeList.Add "Acetate"
	MediaTypeList.Add "Flexi-disc"
	MediaTypeList.Add "Lathe Cut"
	MediaTypeList.Add "Shellac"
	MediaTypeList.Add "Pathé Disc"
	MediaTypeList.Add "Edison Disc"
	MediaTypeList.Add "Cylinder"
	MediaTypeList.Add "CDr"
	MediaTypeList.Add "CDV"
	MediaTypeList.Add "DVDr"
	MediaTypeList.Add "HD DVD"
	MediaTypeList.Add "HD DVD-R"
	MediaTypeList.Add "Blue-ray-R"
	MediaTypeList.Add "4-Track Cartridge"
	MediaTypeList.Add "8-Track Cartridge"
	MediaTypeList.Add "DCC"
	MediaTypeList.Add "Microcassette"
	MediaTypeList.Add "Reel-To-Reel"
	MediaTypeList.Add "Betamax"
	MediaTypeList.Add "VHS"
	MediaTypeList.Add "Video 2000"
	MediaTypeList.Add "Laserdisc"
	MediaTypeList.Add "SelectaVision"
	MediaTypeList.Add "VHD"
	MediaTypeList.Add "MVD"
	MediaTypeList.Add "UMD"
	MediaTypeList.Add "Floppy Disk"
	MediaTypeList.Add "Memory Stick"
	MediaTypeList.Add "Hybrid"
	MediaTypeList.Add "Box Set"
	
	MediaFormatList.Add "None"
	MediaFormatList.Add "Album"
	MediaFormatList.Add "Mini-Album"
	MediaFormatList.Add "Compilation"
	MediaFormatList.Add "Single"
	MediaFormatList.Add "Maxi-Single"
	MediaFormatList.Add "7"""
	MediaFormatList.Add "12"""
	MediaFormatList.Add "LP"
	MediaFormatList.Add "EP"
	MediaFormatList.Add "Single Sided"
	MediaFormatList.Add "Enhanced"
	MediaFormatList.Add "Limited Edition"
	MediaFormatList.Add "Reissue"
	MediaFormatList.Add "Remastered"
	MediaFormatList.Add "Repress"
	MediaFormatList.Add "Test Pressing"
	MediaFormatList.Add "Unofficial"
	MediaFormatList.Add "Promo"
	MediaFormatList.Add "White Label"
	MediaFormatList.Add "Mixed"
	MediaFormatList.Add "Sampler"
	MediaFormatList.Add "MP3"
	MediaFormatList.Add "FLAC"
	MediaFormatList.Add "16"""
	MediaFormatList.Add "11"""
	MediaFormatList.Add "10"""
	MediaFormatList.Add "9"""
	MediaFormatList.Add "8"""
	MediaFormatList.Add "6"""
	MediaFormatList.Add "5"""
	MediaFormatList.Add "4"""
	MediaFormatList.Add "3"""
	MediaFormatList.Add "45 RPM"
	MediaFormatList.Add "78 RPM"
	MediaFormatList.Add "Shape"
	MediaFormatList.Add "Card Backed"
	MediaFormatList.Add "Etched"
	MediaFormatList.Add "Picture Disc"
	MediaFormatList.Add "Stereo"
	MediaFormatList.Add "Mono"
	MediaFormatList.Add "Quadraphonic"
	MediaFormatList.Add "Ambisonic"
	MediaFormatList.Add "Mispress"
	MediaFormatList.Add "Misprint"
	MediaFormatList.Add "Partially Mixed"
	MediaFormatList.Add "Unofficial Release"
	MediaFormatList.Add "Partially Unofficial"
	MediaFormatList.Add "Copy Protected"
	
	CountryList.Add "None"
	CountryList.Add "Australia"
	CountryList.Add "Belgium"
	CountryList.Add "Brazil"
	CountryList.Add "Canada"
	CountryList.Add "China"
	CountryList.Add "Cuba"
	CountryList.Add "France"
	CountryList.Add "Germany"
	CountryList.Add "Italy"
	CountryList.Add "Jamaica"
	CountryList.Add "Japan"
	CountryList.Add "Ireland"
	CountryList.Add "India"
	CountryList.Add "Mexico"
	CountryList.Add "Netherlands"
	CountryList.Add "New Zealand"
	CountryList.Add "Spain"
	CountryList.Add "Sweden"
	CountryList.Add "Switzerland"
	CountryList.Add "UK"
	CountryList.Add "US"
	CountryList.Add "=========="
	CountryList.Add "Africa"
	CountryList.Add "Asia"
	CountryList.Add "Australasia"
	CountryList.Add "Benelux"
	CountryList.Add "Central America"
	CountryList.Add "Europe"
	CountryList.Add "Gulf Cooperation Council"
	CountryList.Add "North America"
	CountryList.Add "Scandinavia"
	CountryList.Add "South America"
	CountryList.Add "==========="
	CountryList.Add "Afghanistan"
	CountryList.Add "Akrotiri"
	CountryList.Add "Albania"
	CountryList.Add "Algeria"
	CountryList.Add "American Samoa"
	CountryList.Add "Andorra"
	CountryList.Add "Angola"
	CountryList.Add "Anguilla"
	CountryList.Add "Antarctica"
	CountryList.Add "Antigua & Barbuda"
	CountryList.Add "Argentina"
	CountryList.Add "Armenia"
	CountryList.Add "Aruba"
	CountryList.Add "Ashmore & Cartier Islands"
	CountryList.Add "Austria"
	CountryList.Add "Azerbaijan"
	CountryList.Add "Bahamas"
	CountryList.Add "Bahrain"
	CountryList.Add "Baker Island"
	CountryList.Add "Bangladesh"
	CountryList.Add "Barbados"
	CountryList.Add "Bassas da India"
	CountryList.Add "Belarus"
	CountryList.Add "Belize"
	CountryList.Add "Benin"
	CountryList.Add "Bermuda"
	CountryList.Add "Bhutan"
	CountryList.Add "Bolivia"
	CountryList.Add "Bosnia & Herzegovina"
	CountryList.Add "Botswana"
	CountryList.Add "Bouvet Island"
	CountryList.Add "British Indian Ocean"
	CountryList.Add "British Virgin Islands"
	CountryList.Add "Brunei"
	CountryList.Add "Bulgaria"
	CountryList.Add "Burkina Faso"
	CountryList.Add "Burma"
	CountryList.Add "Burundi"
	CountryList.Add "Cambodia"
	CountryList.Add "Cameroon"
	CountryList.Add "Cape Verde"
	CountryList.Add "Cayman Islands"
	CountryList.Add "Central African Republic"
	CountryList.Add "Chad"
	CountryList.Add "Chile"
	CountryList.Add "Christmas Island"
	CountryList.Add "Clipperton Island"
	CountryList.Add "Cocos Islands"
	CountryList.Add "Colombia"
	CountryList.Add "Comoros"
	CountryList.Add "Congo"
	CountryList.Add "Cook Islands"
	CountryList.Add "Coral Sea Islands"
	CountryList.Add "Costa Rica"
	CountryList.Add "Croatia"
	CountryList.Add "Cyprus"
	CountryList.Add "Czech Republic"
	CountryList.Add "Czechoslovakia"
	CountryList.Add "Denmark"
	CountryList.Add "Dhekelia"
	CountryList.Add "Djibouti"
	CountryList.Add "Dominica"
	CountryList.Add "Dominican Republic"
	CountryList.Add "East Timor"
	CountryList.Add "Ecuador"
	CountryList.Add "Egypt"
	CountryList.Add "El Salvador"
	CountryList.Add "Equatorial Guinea"
	CountryList.Add "Eritrea"
	CountryList.Add "Estonia"
	CountryList.Add "Ethiopia"
	CountryList.Add "Europa Island"
	CountryList.Add "Falkland Islands"
	CountryList.Add "Faroe Islands"
	CountryList.Add "Fiji"
	CountryList.Add "Finland"
	CountryList.Add "French Guiana"
	CountryList.Add "French Polynesia"
	CountryList.Add "French Southern"
	CountryList.Add "Gabon"
	CountryList.Add "Gambia"
	CountryList.Add "Gaza Strip"
	CountryList.Add "Georgia"
	CountryList.Add "German Democratic Republic"
	CountryList.Add "Ghana"
	CountryList.Add "Gibraltar"
	CountryList.Add "Glorioso Islands"
	CountryList.Add "Greece"
	CountryList.Add "Greenland"
	CountryList.Add "Grenada"
	CountryList.Add "Guadeloupe"
	CountryList.Add "Guam"
	CountryList.Add "Guatemala"
	CountryList.Add "Guernsey"
	CountryList.Add "Guinea"
	CountryList.Add "Guinea-Bissau"
	CountryList.Add "Guyana"
	CountryList.Add "Haiti"
	CountryList.Add "Heard Island"
	CountryList.Add "McDonald Islands"
	CountryList.Add "Holy See"
	CountryList.Add "Honduras"
	CountryList.Add "Hong Kong"
	CountryList.Add "Howland Island"
	CountryList.Add "Hungary"
	CountryList.Add "Iceland"
	CountryList.Add "Indonesia"
	CountryList.Add "Iran"
	CountryList.Add "Iraq"
	CountryList.Add "Israel"
	CountryList.Add "Ivory Coast"
	CountryList.Add "Jan Mayen"
	CountryList.Add "Jarvis Island"
	CountryList.Add "Jersey"
	CountryList.Add "Johnston Atoll"
	CountryList.Add "Jordan"
	CountryList.Add "Juan de Nova Island"
	CountryList.Add "Kazakhstan"
	CountryList.Add "Kenya"
	CountryList.Add "Kingman Reef"
	CountryList.Add "Kiribati"
	CountryList.Add "Kuwait"
	CountryList.Add "Kyrgyzstan"
	CountryList.Add "Laos"
	CountryList.Add "Latvia"
	CountryList.Add "Lebanon"
	CountryList.Add "Lesotho"
	CountryList.Add "Liberia"
	CountryList.Add "Libya"
	CountryList.Add "Liechtenstein"
	CountryList.Add "Lithuania"
	CountryList.Add "Luxembourg"
	CountryList.Add "Macau"
	CountryList.Add "Macedonia"
	CountryList.Add "Madagascar"
	CountryList.Add "Malawi"
	CountryList.Add "Malaysia"
	CountryList.Add "Maldives"
	CountryList.Add "Mali"
	CountryList.Add "Malta"
	CountryList.Add "Man, Isle of"
	CountryList.Add "Marshall Islands"
	CountryList.Add "Martinique"
	CountryList.Add "Mauritania"
	CountryList.Add "Mauritius"
	CountryList.Add "Mayotte"
	CountryList.Add "Micronesia"
	CountryList.Add "Midway Islands"
	CountryList.Add "Moldova"
	CountryList.Add "Monaco"
	CountryList.Add "Mongolia"
	CountryList.Add "Montenegro"
	CountryList.Add "Montserrat"
	CountryList.Add "Morocco"
	CountryList.Add "Mozambique"
	CountryList.Add "Namibia"
	CountryList.Add "Nauru"
	CountryList.Add "Navassa Island"
	CountryList.Add "Nepal"
	CountryList.Add "Netherlands Antilles"
	CountryList.Add "New Caledonia"
	CountryList.Add "Nicaragua"
	CountryList.Add "Niger"
	CountryList.Add "Nigeria"
	CountryList.Add "Niue"
	CountryList.Add "Norfolk Island"
	CountryList.Add "Northern Mariana Islands"
	CountryList.Add "North Korea"
	CountryList.Add "Norway"
	CountryList.Add "Oman"
	CountryList.Add "Pakistan"
	CountryList.Add "Palau"
	CountryList.Add "Palmyra Atoll"
	CountryList.Add "Panama"
	CountryList.Add "Papua New Guinea"
	CountryList.Add "Paracel Islands"
	CountryList.Add "Paraguay"
	CountryList.Add "Peru"
	CountryList.Add "Philippines"
	CountryList.Add "Pitcairn Islands"
	CountryList.Add "Poland"
	CountryList.Add "Portugal"
	CountryList.Add "Puerto Rico"
	CountryList.Add "Qatar"
	CountryList.Add "Reunion"
	CountryList.Add "Romania"
	CountryList.Add "Russia"
	CountryList.Add "Rwanda"
	CountryList.Add "Saint Helena"
	CountryList.Add "Saint Kitts and Nevis"
	CountryList.Add "Saint Lucia"
	CountryList.Add "Saint Pierre"
	CountryList.Add "Saint Vincent"
	CountryList.Add "Samoa"
	CountryList.Add "San Marino"
	CountryList.Add "Sao Tome and Principe"
	CountryList.Add "Saudi Arabia"
	CountryList.Add "Senegal"
	CountryList.Add "Serbia"
	CountryList.Add "Serbia and Montenegro"
	CountryList.Add "Seychelles"
	CountryList.Add "Sierra Leone"
	CountryList.Add "Singapore"
	CountryList.Add "Slovakia"
	CountryList.Add "Slovenia"
	CountryList.Add "Solomon Islands"
	CountryList.Add "Somalia"
	CountryList.Add "South Africa"
	CountryList.Add "South Korea"
	CountryList.Add "Spratly Islands"
	CountryList.Add "Sri Lanka"
	CountryList.Add "Sudan"
	CountryList.Add "Suriname"
	CountryList.Add "Svalbard"
	CountryList.Add "Swaziland"
	CountryList.Add "Syria"
	CountryList.Add "Tajikistan"
	CountryList.Add "Tanzania"
	CountryList.Add "Thailand"
	CountryList.Add "Taiwan"
	CountryList.Add "Togo"
	CountryList.Add "Tokelau"
	CountryList.Add "Tonga"
	CountryList.Add "Trinidad & Tobago"
	CountryList.Add "Tromelin Island"
	CountryList.Add "Tunisia"
	CountryList.Add "Turkey"
	CountryList.Add "Turkmenistan"
	CountryList.Add "Turks and Caicos Islands"
	CountryList.Add "Tuvalu"
	CountryList.Add "Uganda"
	CountryList.Add "Ukraine"
	CountryList.Add "United Arab Emirates"
	CountryList.Add "Uruguay"
	CountryList.Add "USSR"
	CountryList.Add "Uzbekistan"
	CountryList.Add "Vatican City"
	CountryList.Add "Vanuatu"
	CountryList.Add "Venezuela"
	CountryList.Add "Vietnam"
	CountryList.Add "Virgin Islands"
	CountryList.Add "Wake Island"
	CountryList.Add "Wallis and Futuna"
	CountryList.Add "West Bank"
	CountryList.Add "Western Sahara"
	CountryList.Add "Yemen"
	CountryList.Add "Yugoslavia"
	CountryList.Add "Zambia"
	CountryList.Add "Zimbabwe"
	
	YearList.Add "None"
	For i=Year(Date) To 1900 Step -1
		YearList.Add i
	Next
	
	If UBound(tmpYear2) <> YearList.Count -1 Then
		'MsgBox UBound(tmpYear2) & " -- " & YearList.Count -1
		tmpYear = tmpYear & ",1"
		ini.StringValue("DiscogsAutoTagWeb","CurrentYearFilter") = tmpYear
		tmpYear2 = Split(tmpYear, ",")
	End If
	
	For a = 0 To CountryList.Count - 1
		CountryFilterList.Add tmpCountry2(a)
	Next
	
	For a = 0 To MediaTypeList.Count - 1
		MediaTypeFilterList.Add tmpMediaType2(a)
	Next
	
	For a = 0 To MediaFormatList.Count - 1
		MediaFormatFilterList.Add tmpMediaFormat2(a)
	Next
	
	For a = 0 To YearList.Count - 1
		YearFilterList.Add tmpYear2(a)
	Next
	
	If MediaTypeFilterList.Item(0) = "0" Then
		FilterMediaType = "None"
	ElseIf MediaTypeFilterList.Item(0) = "1" Then
		FilterMediaType = "Use MediaType Filter"
	Else
		FilterMediaType = MediaTypeFilterList.Item(0)
	End If
	
	If MediaFormatFilterList.Item(0) = "0" Then
		FilterMediaFormat = "None"
	ElseIf MediaFormatFilterList.Item(0) = "1" Then
		FilterMediaFormat = "Use MediaFormat Filter"
	Else
		FilterMediaFormat = MediaFormatFilterList.Item(0)
	End If
	
	If CountryFilterList.Item(0) = "0" Then
		FilterCountry = "None"
	ElseIf CountryFilterList.Item(0) = "1" Then
		FilterCountry = "Use Country Filter"
	Else
		FilterCountry = CountryFilterList.Item(0)
	End If
	
	If YearFilterList.Item(0) = "0" Then
		FilterYear = "None"
	ElseIf YearFilterList.Item(0) = "1" Then
		FilterYear = "Use Year Filter"
	Else
		FilterYear = YearFilterList.Item(0)
	End If
	
	
	Rem Dim ErrorMessage
	Dim ret, res
	Dim itm, found
	Dim LAlbumTracks
	
	
	
	'SongList --> Alle ausgewählten Songs aus Mediamonkey
	'AnzahlSongs --> Anzahl der ausgewählten Songs aus MM

	'AlbumArtist_AlbumList --> Liste der unterschiedlichen Alben (AlbumArtistName AlbumName)
	'SongIDList --> ID des ersten Songs des Albums
	'  --> ID des Albums
	'Beispiel:
	'AlbumArtist_AlbumList.Item(0) = "Metallica Ride the Lightining"
	'SongIDList.Item(0) = 15 (der 16. Song in der ausgewählten Liste aus MM)
	
	'NewTrackList --> Die Songs des zu bearbeitenden Album
	
	
	Set SongList = SDB.SelectedSongList
	If SongList.count = 0 Then
		Set SongList = SDB.AllVisibleSongList
	End If
	If SongList.count = 0 Then
		ErrorMessage = "No Songs selected"
		FormatErrorMessage ErrorMessage
	End If
	
	Set AlbumIDList = SDB.NewStringList
	
	WriteLog "Start song analysis"
	WriteLog "Songs count=" & SongList.count
	
	For i = 0 To SongList.Count - 1
		Set itm = SongList.Item(i)
		WriteLog "Song " & i+1
		Dim itmAlbum : Set itmAlbum = itm.Album
		WriteLog "AlbumID=" & itmAlbum.ID & "  /  Artist=" & itm.ArtistName & "  /  Title=" & itm.Title & "  /  Album=" & itm.AlbumName
		Dim LAlbumName : LAlbumName = itm.AlbumName
		Dim LAlbumArtistName : LAlbumArtistName = itm.AlbumArtistName
		Dim LAlbumID : LAlbumID = itmAlbum.ID
		Dim cnt
		
		If LAlbumID <> -1 Then
			If AlbumIDList.Count > 0 Then
				found = 0
				For cnt = 0 To AlbumIDList.Count -1
					If CLng(AlbumIDList.Item(cnt)) =  LAlbumID Then
						found = 1
					End If
				Next
				
				If found = 0 Then
					AlbumIDList.Add LAlbumID
				End If
				
			Else
				AlbumIDList.Add LAlbumID
			End If
		End If
	Next
	
	
	CurrentSelectedAlbum = 0
	WriteLog "AlbumIDList.Count (number of different albums)=" & AlbumIDList.Count
	If AlbumIDList.Count = 0 Then
		res = SDB.MessageBox("You selected no songs with filled album tag. The script now exit", mtInformation, Array(mbOk))
		Set ini = Nothing
		Set ResultsReleaseID = Nothing
		Script.UnregisterAllEvents
		Exit Sub
	End If
	
	Rem 'Create the window to be shown
	Dim BottomF
	Set Form = UI.NewForm
	Form.FormPosition = 4
	Form.Caption = "Options for Discogs Search"
	'FormBorderStyle = 0
	Form.Common.ClientWidth = 280
	Form.Common.ClientHeight = 250
	Form.StayOnTop = True
	Set BottomF = UI.NewPanel(Form)
	BottomF.Common.Align = 2   ' Bottom
	BottomF.Common.Height = 30
	
	Dim OptionsHTML, OptionsHTMLDoc
	OptionsHTML= "<HTML>"
	OptionsHTML = OptionsHTML &  "<HEAD>"
	OptionsHTML = OptionsHTML &  "<style type=""text/css"" media=""screen"">"
	OptionsHTML = OptionsHTML &  ".tabletext { font-family: Arial, Helvetica, sans-serif; font-size: 8pt;}"
	OptionsHTML = OptionsHTML &  "</style>"
	OptionsHTML = OptionsHTML &  "</HEAD>"
	OptionsHTML = OptionsHTML &  "<body bgcolor=""#FFFFFF"">"
	OptionsHTML = OptionsHTML &  "<table border=0 width=100% cellspacing=0 cellpadding=1 class=tabletext>"
	OptionsHTML = OptionsHTML &  "<tr>"
	OptionsHTML = OptionsHTML &  "<td align=left><b>DiscogsAutoTagWeb Batch Version " & VersionStr & "</b></td>"
	OptionsHTML = OptionsHTML &  "<td colspan=3 align=right valign=top>"
	OptionsHTML = OptionsHTML &  "<tr><td colspan=2 align=center><br></td></tr>"
	OptionsHTML = OptionsHTML &  "<tr><td colspan=2 align=center>You selected " & AlbumIDList.Count & " Albums</td></tr>"
	OptionsHTML = OptionsHTML &  "<tr><td colspan=2 align=center>The script now searches on discogs</td></tr>"
	OptionsHTML = OptionsHTML &  "<tr><td colspan=2 align=center><br></td></tr>"
	OptionsHTML = OptionsHTML &  "<tr><td align=center colspan=2><b>Options:</b></td></tr>"
	OptionsHTML = OptionsHTML &  "<tr><td colspan=2 align=left><input type=checkbox id=""SkipNotChangedReleases"" title=""Skip unchanged releases automatically and check the next one"" >Skip unchanged releases</td></tr>"
	OptionsHTML = OptionsHTML &  "<tr><td colspan=2 align=left><input type=checkbox id=""ProcessOnlyDiscogs"" title=""Process only albums already found at discogs"" >Process only Discogs Releases</td></tr>"
	OptionsHTML = OptionsHTML &  "<tr><td colspan=2 align=left><input type=checkbox id=""ProcessNoDiscogs"" title=""Process only albums not found at discogs"" >Process no Discogs releases</td></tr>"
	OptionsHTML = OptionsHTML &  "<tr><td colspan=2 align=center><br></td></tr>"
	
	OptionsHTML = OptionsHTML &  "</table>"
	OptionsHTML = OptionsHTML &  "</td>"
	OptionsHTML = OptionsHTML & "</table>"
	OptionsHTML = OptionsHTML &  "</body>"
	OptionsHTML = OptionsHTML &  "</HTML>"
	
	Dim WebBrowser3
	Set WebBrowser3 = UI.NewActiveX(Form, "Shell.Explorer")
	WebBrowser3.Common.Align = 0      ' Fill whole client rectangle
	WebBrowser3.Common.ControlName = "WebBrowser3"
	WebBrowser3.Common.Top = 0
	WebBrowser3.Common.Left = 0
	WebBrowser3.Common.Height = 220
	WebBrowser3.Common.Width = 280
	SDB.Objects("WebBrowser3") = WebBrowser3
	WebBrowser3.Interf.Visible = True
	WebBrowser3.Common.BringToFront
	WebBrowser3.SetHTMLDocument OptionsHTML
	
	Set OptionsHTMLDoc = WebBrowser3.Interf.Document
	
	Dim Btn8 : Set Btn8 = UI.NewButton(BottomF)
	Btn8.Common.ControlName = "Btn8"
	Btn8.Common.SetRect 30, 5, 80, 20
	Btn8.Caption = SDB.Localize("Ok")
	Btn8.Common.Anchors = 6
	Btn8.ModalResult = 2
	
	Dim Btn9 : Set Btn9 = UI.NewButton(BottomF)
	Btn9.Common.ControlName = "Btn9"
	Btn9.Common.SetRect 140, 5, 80, 20
	Btn9.Caption = SDB.Localize("Cancel")
	Btn9.Common.Hint = "Discard the changes"
	Btn9.Common.Anchors = 6
	Btn9.ModalResult = 1
	
	Dim checkbox
	Set checkBox = OptionsHTMLDoc.getElementById("ProcessOnlyDiscogs")
	checkBox.checked = ProcessOnlyDiscogs

	Set checkBox = OptionsHTMLDoc.getElementById("ProcessNoDiscogs")
	checkBox.checked = ProcessNoDiscogs

	Set checkBox = OptionsHTMLDoc.getElementById("SkipNotChangedReleases")
	checkBox.Checked = SkipNotChangedReleases

	If Not (Form.ShowModal = 2) Then
		FinishSearch Form
		Exit Sub
	Else
		If Not (ini Is Nothing) Then
			Set checkBox = OptionsHTMLDoc.getElementById("SkipNotChangedReleases")
			SkipNotChangedReleases = checkBox.Checked
			ini.BoolValue("DiscogsAutoTagWeb","SkipNotChangedReleases") = SkipNotChangedReleases
			Set checkBox = OptionsHTMLDoc.getElementById("ProcessOnlyDiscogs")
			ProcessOnlyDiscogs = checkBox.checked
			ini.BoolValue("DiscogsAutoTagWeb","ProcessOnlyDiscogs") = ProcessOnlyDiscogs
			Set checkBox = OptionsHTMLDoc.getElementById("ProcessNoDiscogs")
			ProcessNoDiscogs = checkBox.checked
			ini.BoolValue("DiscogsAutoTagWeb","ProcessNoDiscogs") = ProcessNoDiscogs
		End If

		RadioBoxCheck = -1

		SDB.ProcessMessages

		NewSearch CurrentSelectedAlbum
		WriteLog "Stop BatchDiscogsSearch"
	End If
End Sub

Sub Update_SkipNotChangedReleases()

	Set WebBrowser3 = SDB.Objects("WebBrowser3")
	Set OptionsHTMLDoc = WebBrowser3.Interf.Document
	Set checkBox = OptionsHTMLDoc.getElementById("SkipNotChangedReleases")
	SkipNotChangedReleases = checkBox.Checked

End Sub


Sub FindResults(SearchTerm, SearchArtist, SearchAlbum)

	WriteLog "Start FindResults"

	Dim searchURL, FilterFound
	Set Combo = Nothing
	Set Combo = UI.NewDropDown(Head)
	Combo.Common.SetRect 5, 5, SearchFormWidth -550, 20
	Combo.Common.Anchors = 6
	Combo.Style = 2     ' List
	Script.RegisterEvent Combo, "OnSelect", "ComboChange"
	Set ResultsReleaseID = SDB.NewStringList
	ErrorMessage = ""

	If (InStr(SearchTerm," - [search by release id]") > 0) Then
		SearchTerm = Left(SearchTerm,InStrRev(SearchTerm," - [search by release id]")-1)
	End If

	If (InStr(SearchTerm," - [search by release url]") > 0) Then
		SearchTerm = Left(SearchTerm,InStrRev(SearchTerm," - [search by release url]")-1)
	End If

	If (InStr(SearchTerm," - [currently tagged with this release]") > 0) Then
		SearchTerm = Left(SearchTerm,InStrRev(SearchTerm," - [currently tagged with this release]")-1)
	End If
	
	If (InStr(SearchTerm," - [search returned no results]") > 0) Then
		SearchTerm = Left(SearchTerm,InStrRev(SearchTerm," - [search returned no results]")-1)
	End If
	
	If (InStr(SearchTerm," - [search that yielded error]") > 0) Then
		SearchTerm = Left(SearchTerm,InStrRev(SearchTerm," - [search that yielded error]")-1)
	End If
	
	' Handle direct urls
	
	If (InStr(SearchTerm,"/master/") > 0) Then
		CurrentLoadType = "Master Release"
		WriteLog "Loadtype Master-Release"
		LoadMasterResults Mid(SearchTerm,InStrRev(SearchTerm,"/")+1)
		Exit Sub
	End If
	
	'Will not be longer supported, cause the artist url at Discogs have no more artist-id
	Rem If (InStr(SearchTerm,"/artist/") > 0) Then
	Rem CurrentLoadType = "Releases of Artist"
	Rem WriteLog "Loadtype Releases of Artist"
	Rem LoadArtistResults Mid(SearchTerm,InstrRev(SearchTerm,"/")+1)
	Rem Exit Sub
	Rem End If
	
	If (InStr(SearchTerm,"/label/") > 0) Then
		CurrentLoadType = "Releases of Label"
		WriteLog "Releases of Label"
		LoadLabelResults Mid(SearchTerm,InStrRev(SearchTerm,"/")+1)
		Exit Sub
	End If
	
	If SearchTerm = "" Then
		ErrorMessage = "No search term"
	ElseIf IsNumeric(SearchTerm) Then
		Combo.AddItem SearchTerm & " - [search by release id]"
	ElseIf (InStr(SearchTerm,"/release/") > 0) Then
		Combo.AddItem SearchTerm & " - [search by release url]"
		ResultsReleaseID.Add Mid(SearchTerm,InStrRev(SearchTerm,"/")+1)
	Else
		If IsNumeric(SavedReleaseId) And Not SavedReleaseId = "" Then
			WriteLog "SavedReleaseID found"
			Set FirstTrack = NewTrackList.item(0)
			Combo.AddItem FirstTrack.Artist.Name & " - " & FirstTrack.Album.Name & " - [currently tagged with this release]"
			ResultsReleaseID.Add get_release_ID(FirstTrack)
		End If

		CurrentLoadType = "Search Results"

		WriteLog "FindResults SearchTerm=" & SearchTerm
		WriteLog "FindResults SearchArtist=" & SearchArtist
		WriteLog "FindResults SearchAlbum=" & SearchAlbum

		If SearchArtist <> "" And SearchAlbum <> "" Then
			searchURL = "http://api.discogs.com/database/search?q=" & URLEncodeUTF8(CleanSearchString(SearchTerm)) & "&type=release&per_page=100"
		ElseIf SearchArtist = "" And SearchAlbum <> "" Then
			searchURL = "http://api.discogs.com/database/search?type=release&title=" & URLEncodeUTF8(CleanSearchString(SearchAlbum)) & "&per_page=100"
		ElseIf SearchArtist <> "" And SearchAlbum = "" Then
			searchURL = "http://api.discogs.com/database/search?type=release&artist=" & URLEncodeUTF8(CleanSearchString(SearchArtist)) & "&per_page=100"
		Else
			searchURL = "http://api.discogs.com/database/search?q=" & URLEncodeUTF8(CleanSearchString(SearchTerm)) & "&type=release&per_page=100"
		End If

		JSONParser_find_result searchURL, "results"
		
		WriteLog "JSON-Parser done"
		
		If ResultsReleaseID.Count = 0 Then
			FilterFound = False
			If FilterCountry = "Use Country Filter" Then
				For a = 1 To CountryList.Count - 1
					If CountryFilterList.Item(a) = "1" Then
						FilterFound = True
						Exit For
					End If
				Next
				If FilterFound = False Then
					ErrorMessage = "No Country Filter set !"
				Else
					ErrorMessage = "Search returned no results"
				End If
				Combo.AddItem SearchTerm & " - [search returned no results]"
			End If
			FilterFound = False
			If FilterMediaType = "Use MediaType Filter" Then
				For a = 1 To MediaTypeList.Count - 1
					If MediaTypeFilterList.Item(a) = "1" Then
						FilterFound = True
						Exit For
					End If
				Next
				If FilterFound = False Then
					If ErrorMessage = "" Then
						ErrorMessage = "No MediaType Filter set !"
					Else
						ErrorMessage = ErrorMessage & vbCrLf & "No MediaType Filter set !"
					End If
				End If
			End If
			FilterFound = False
			If FilterMediaFormat = "Use MediaFormat Filter" Then
				For a = 1 To MediaFormatList.Count - 1
					If MediaFormatFilterList.Item(a) = "1" Then
						FilterFound = True
						Exit For
					End If
				Next
				If FilterFound = False Then
					If ErrorMessage = "" Then
						ErrorMessage = "No MediaFormat Filter set !"
					Else
						ErrorMessage = ErrorMessage & vbCrLf & "No MediaFormat Filter set !"
					End If
				End If
			End If
			FilterFound = False
			If FilterYear = "Use Year Filter" Then
				For a = 1 To YearList.Count - 1
					If YearFilterList.Item(a) = "1" Then
						FilterFound = True
						Exit For
					End If
				Next
				If FilterFound = False Then
					If ErrorMessage = "" Then
						ErrorMessage = "No Year Filter set !"
					Else
						ErrorMessage = ErrorMessage & vbCrLf & "No Year Filter set !"
					End If
				End If
			End If
			
			If ErrorMessage = "" Then
				ErrorMessage = "Search returned no results"
			End If
			Combo.AddItem SearchTerm & " - [search returned no results]"
			WriteLog("SearchTerm = " & SearchTerm)
		End If
		
	End If
	
	Combo.ItemIndex = 0
'	ShowResult 0
	WriteLog "Stop FindResults"
	
End Sub


Sub LoadMasterResults(MasterId)
	
	Dim masterURL
	WriteLog "Start LoadMasterResults"
	
	Set Combo = Nothing
	Set Combo = UI.NewDropDown(Head)
	Combo.Common.SetRect 5, 5, SearchFormWidth -550, 20
	Combo.Common.Anchors = 6
	Combo.Style = 2     ' List
	Script.RegisterEvent Combo, "OnSelect", "ComboChange"
	
	Set ResultsReleaseID = SDB.NewStringList
	ErrorMessage = ""
	
	If MasterId = "" Then
		ErrorMessage = "Cannot load empty master release"
	Else
		If IsNumeric(SavedReleaseId) And Not SavedReleaseId = "" Then
			Set FirstTrack = NewTrackList.item(0)
			Combo.AddItem FirstTrack.Artist.Name & " - " & FirstTrack.Album.Name & " - [currently tagged with this release]"
			ResultsReleaseID.Add get_release_ID(FirstTrack)
		End If
		
		masterURL = "http://api.discogs.com/masters/" & MasterId & "/versions"
		JSONParser_find_result masterURL, "versions"
	End If
	
	If ErrorMessage <> "" Then
		FormatErrorMessage ErrorMessage
	Else
		Combo.ItemIndex = 0
		ShowResult 0
	End If
	
	WriteLog "Stop LoadMasterResults"
	
End Sub


Sub LoadArtistResults(ArtistId)

	Dim artistURL
	WriteLog "Start LoadArtistResults"
	
	Set Combo = Nothing
	Set Combo = UI.NewDropDown(Head)
	Combo.Common.SetRect 5, 5, SearchFormWidth -550, 20
	Combo.Common.Anchors = 6
	Combo.Style = 2     ' List
	Script.RegisterEvent Combo, "OnSelect", "ComboChange"
	
	Set ResultsReleaseID = SDB.NewStringList
	ErrorMessage = ""
	
	If ArtistId = "" Then
		ErrorMessage = "Cannot load empty artist"
	Else
		If IsNumeric(SavedReleaseId) And Not SavedReleaseId = "" Then
			Set FirstTrack = NewTrackList.item(0)
			Combo.AddItem FirstTrack.Artist.Name & " - " & FirstTrack.Album.Name & " - [currently tagged with this release]"
			ResultsReleaseID.Add get_release_ID(FirstTrack)
		End If
		
		artistURL = "http://api.discogs.com/artists/" & ArtistId & "/releases&per_page=100"
		WriteLog "ArtistSuchURL=" & artistURL
		JSONParser_find_result artistURL, "releases"
	End If
	
	If ErrorMessage <> "" Then
		FormatErrorMessage ErrorMessage
	Else
		Combo.ItemIndex = 0
		ShowResult 0
	End If
	
	WriteLog "Stop LoadArtistResults"
	
End Sub


Sub LoadLabelResults(LabelId)
	
	Dim labelURL
	WriteLog "Start LoadLabelResults"
	
	Set Combo = Nothing
	Set Combo = UI.NewDropDown(Head)
	Combo.Common.SetRect 5, 5, SearchFormWidth -550, 20
	Combo.Common.Anchors = 6
	Combo.Style = 2     ' List
	Script.RegisterEvent Combo, "OnSelect", "ComboChange"
	
	Set ResultsReleaseID = SDB.NewStringList
	ErrorMessage = ""
	
	If LabelId = "" Then
		ErrorMessage = "Cannot load empty label"
	Else
		If IsNumeric(SavedReleaseId) And Not SavedReleaseId = "" Then
			Set FirstTrack = NewTrackList.item(0)
			Combo.AddItem FirstTrack.Artist.Name & " - " & FirstTrack.Album.Name & " - [currently tagged with this release]"
			ResultsReleaseID.Add get_release_ID(FirstTrack)
		End If
		
		labelURL = "http://api.discogs.com/labels/" & LabelId & "/releases&per_page=100"
		WriteLog "labelURL=" & labelURL
		JSONParser_find_result labelURL, "releases"
	End If
	
	If ErrorMessage <> "" Then
		FormatErrorMessage ErrorMessage
	Else
		Combo.ItemIndex = 0
		ShowResult 0
	End If
	
	WriteLog "Stop LoadLabelResults"
	
End Sub


'For reloading results
Sub ReloadResults

	WriteLog "Start ReloadResults"

	Dim DiscogsTracksNum, Durations
	Dim AlbumLyricist, AlbumComposer, AlbumConductor, AlbumProducer, AlbumInvolved, AlbumFeaturing
	Dim NoSubTrackUsing

	Set Tracks = SDB.NewStringList
	Set TracksNum = SDB.NewStringList
	Set DiscogsTracksNum = SDB.NewStringList
	Set TracksCD = SDB.NewStringList
	Set ArtistTitles = SDB.NewStringList
	Set InvolvedArtists = SDB.NewStringList
	Set Lyricists = SDB.NewStringList
	Set Composers = SDB.NewStringList
	Set Conductors = SDB.NewStringList
	Set Producers = SDB.NewStringList
	Set Durations = SDB.NewStringList
	
	'----------------------------------DiscogsImages----------------------------------------
	Rem Set SaveImage = SDB.NewStringList
	Rem Set SaveImageType = SDB.NewStringList
	Rem Set FileNameList = SDB.NewStringList
	Rem ImagesCount = 0
	'----------------------------------DiscogsImages----------------------------------------
	
	If Not IsNull(CurrentRelease) Then
		
		AlbumArtist = ""
		AlbumArtistTitle = ""
		AlbumLyricist = ""
		AlbumComposer = ""
		AlbumConductor = ""
		AlbumProducer = ""
		AlbumInvolved = ""
		AlbumArtURL = ""
		AlbumArtThumbNail = ""
		AlbumFeaturing = ""
		LastDisc = ""
		
		Dim iTrackNum, iSubTrack, cSubTrack, subTrackTitle
		Dim trackName, t, pos
		Dim role, rolea, currentRole, NoSplit, zahl, zahltemp, zahl2, zahltemp2
		Dim CharSeparatorSubTrack
		ReDim Involved_R(0)
		Dim tmp
		Dim tmp2
		Dim rTrack
		Dim LeadingZeroTrackPosition
		ReDim TrackRoles(0)
		ReDim TrackArtist2(0)
		ReDim TrackPos(0)
		ReDim Title_Position(0)
		SavedArtistId = ""
		SavedLabelId = ""
		LeadingZeroTrackPosition = False
		
		theLabels = ""
		theFormat = ""
		theCatalogs = ""
		Genres = ""
		Styles = ""
		WebTrackCount = 0
		
		'Get Track-List
		Dim track, currentTrack, position
		For Each track In CurrentRelease("tracklist")
			Set currentTrack = CurrentRelease("tracklist")(track)
			position = currentTrack("position")
			DiscogsTracksNum.Add position
			position = exchange_roman_numbers(position)
			ReDim Preserve Title_Position(UBound(Title_Position)+1)
			Title_Position(UBound(Title_Position)) = position
			WebTrackCount = WebTrackCount + 1
		Next
		
		'Check for leading zero in track-position
		LeadingZeroTrackPosition = CheckLeadingZeroTrackPosition(Title_Position(1))
		
		
		' Get artist title
		Dim artist, currentArtist, artistName
		For Each artist In CurrentRelease("artists")
			Set currentArtist = CurrentRelease("artists")(artist)
			If Not CheckUseAnv And currentArtist("anv") <> "" Then
				artistName = CleanArtistName(currentArtist("anv"))
				' !!!!!artistName <- currentArtist
			Else
				artistName = CleanArtistName(currentArtist("name"))
				' !!!!!artistName <- currentArtist
			End If
			If SavedArtistId = "" Then SavedArtistId = currentArtist("id")
			
			If (AlbumArtist = "") Then
				AlbumArtist = artistName
			End If
			
			Writelog("SavedArtistId=" & SavedArtistId)
			AlbumArtistTitle = AlbumArtistTitle & artistName
			
			If currentArtist("join") <> "" Then
				tmp = currentArtist("join")
				If tmp = "," Then
					AlbumArtistTitle = AlbumArtistTitle & ArtistSeparator
				ElseIf LookForFeaturing(tmp) And CheckFeaturingName Then
					If TxtFeaturingName = "," or TxtFeaturingName = ";" Then
						AlbumArtistTitle = AlbumArtistTitle & TxtFeaturingName & " "
					Else
						AlbumArtistTitle = AlbumArtistTitle & " " & TxtFeaturingName & " "
					End If
				Else
					AlbumArtistTitle = AlbumArtistTitle & " " & currentArtist("join") & " "
				End If
			End If
		Next
		Writelog("AlbumArtistTitle=" & AlbumArtistTitle)

		If Right(AlbumArtistTitle, 3) = " , " Then AlbumArtistTitle = Left(AlbumArtistTitle, Len(AlbumArtistTitle)-3)

		If (Not CheckAlbumArtistFirst) Then
			AlbumArtist = AlbumArtistTitle
		End If
		
		If AlbumArtist = "Various" And CheckVarious Then
			AlbumArtist = TxtVarious
		End If
		If AlbumArtistTitle = "Various" And CheckVarious Then
			AlbumArtistTitle = TxtVarious
		End If
		
		
		WriteLog "ExtraArtists"
		Dim extraArtist
		If currentRelease.Exists("extraartists") Then
			For Each extraArtist In CurrentRelease("extraartists")
				Set currentArtist = CurrentRelease("extraartists")(extraArtist)
				If currentArtist("tracks") = "" Then
					If (currentArtist("anv") <> "") And Not CheckUseAnv Then
						artistName = CleanArtistName(currentArtist("anv"))
					Else
						artistName = CleanArtistName(currentArtist("name"))
					End If
					role = currentArtist("role")
					NoSplit = False
					If InStr(role, ",") = 0 Then
						currentRole = role
						zahl = 0
						NoSplit = True
					Else
						rolea = Split(role, ", ")
						zahl = UBound(rolea)
					End If
					
					For zahltemp = 0 To zahl
						If NoSplit = False Then
							currentRole = rolea(zahltemp)
						End If
						If LookForFeaturing(currentRole) Then
							If InStr(AlbumFeaturing, artistName) = 0 Then
								If AlbumFeaturing = "" Then
									If CheckFeaturingName Then
										AlbumFeaturing = TxtFeaturingName & " " & artistName
									Else
										AlbumFeaturing = currentRole & " " & artistName
									End If
								Else
									AlbumFeaturing = AlbumFeaturing & Separator & artistName
								End If
							End If
						Else
							Do
								tmp = searchKeyword(LyricistKeywords, currentRole, AlbumLyricist, artistName)
								If tmp <> "" And tmp <> "ALREADY_INSIDE_ROLE" Then
									AlbumLyricist = tmp
									Exit Do
								ElseIf tmp = "ALREADY_INSIDE_ROLE" Then
									Exit Do
								End If
								tmp = searchKeyword(ConductorKeywords, currentRole, AlbumConductor, artistName)
								If tmp <> "" And tmp <> "ALREADY_INSIDE_ROLE" Then
									AlbumConductor = tmp
									Exit Do
								ElseIf tmp = "ALREADY_INSIDE_ROLE" Then
									Exit Do
								End If
								tmp = searchKeyword(ProducerKeywords, currentRole, AlbumProducer, artistName)
								If tmp <> "" And tmp <> "ALREADY_INSIDE_ROLE" Then
									AlbumProducer = tmp
									Exit Do
								ElseIf tmp = "ALREADY_INSIDE_ROLE" Then
									Exit Do
								End If
								tmp = searchKeyword(ComposerKeywords, currentRole, AlbumComposer, artistName)
								If tmp <> "" And tmp <> "ALREADY_INSIDE_ROLE" Then
									AlbumComposer = tmp
									Exit Do
								ElseIf tmp = "ALREADY_INSIDE_ROLE" Then
									Exit Do
								End If
								tmp2 = search_involved(Involved_R, currentRole)
								If tmp2 = -1 Then
									ReDim Preserve Involved_R(UBound(Involved_R)+1)
									Involved_R(UBound(Involved_R)) = currentRole & ": " & artistName
								Else
									If InStr(Involved_R(tmp2), artistName) = 0 Then
										Involved_R(tmp2) = Involved_R(tmp2) & ", " & artistName
									End If
								End If
								Exit Do
							Loop While True
						End If
					Next
				Else
					If Not CheckUseAnv And currentArtist("anv") <> "" Then
						artistName = CleanArtistName(currentArtist("anv"))
					Else
						artistName = CleanArtistName(currentArtist("name"))
					End If
					role = currentArtist("role")
					rTrack = currentArtist("tracks")
					NoSplit = False
					If InStr(role, ",") <> 0 Then
						rolea = Split(role, ", ")
						zahl = UBound(rolea)
					ElseIf InStr(role, "&") <> 0 Then
						rolea = Split(role, " & ")
						zahl = UBound(rolea)
					Else
						involvedRole = role
						zahl = 0
						NoSplit = True
					End If
					For zahltemp = 0 To zahl
						If NoSplit = False Then
							involvedRole = rolea(zahltemp)
						End If
						
						If InStr(rTrack, ",") = 0 And InStr(rTrack, " to ") = 0 And InStr(rTrack, " & ") = 0 Then
							currentTrack = rTrack
							Add_Track_Role currentTrack, artistName, involvedRole, TrackRoles, TrackArtist2, TrackPos
						End If
						If InStr(rTrack, ",") <> 0 Then
							tmp = Split(rTrack, ",")
							zahl2 = UBound(tmp)
							For zahltemp2 = 0 To zahl2
								currentTrack = Trim(tmp(zahltemp2))
								If InStr(currentTrack, " to ") <> 0 Then
									Track_from_to currentTrack, artistName, involvedRole, Title_Position, TrackRoles, TrackArtist2, TrackPos, LeadingZeroTrackPosition
								Else
									Add_Track_Role currentTrack, artistName, involvedRole, TrackRoles, TrackArtist2, TrackPos
								End If
							Next
						ElseIf InStr(rTrack, " to ") <> 0 Then
							currentTrack = rTrack
							Track_from_to currentTrack, artistName, involvedRole, Title_Position, TrackRoles, TrackArtist2, TrackPos, LeadingZeroTrackPosition
						ElseIf InStr(rTrack, " & ") <> 0 Then
							tmp = Split(rTrack, " & ")
							zahl2 = UBound(tmp)
							For zahltemp2 = 0 To zahl2
								currentTrack = Trim(tmp(zahltemp2))
								Add_Track_Role currentTrack, artistName, involvedRole, TrackRoles, TrackArtist2, TrackPos
							Next
						End If
					Next
				End If
			Next
		End If
		' Get track titles and track artists
		
		iAutoTrackNumber = 1
		iAutoDiscNumber = 1
		iTrackNum = 0
		iSubTrack = 0
		cSubTrack = -1
		subTrackTitle = ""
		CharSeparatorSubTrack = 0
		Rem CharSeparatorSubTrack: 0 = nothing    1 = "."     2 = a-z
		Rem subTrackStart = 1 '0 = Song -1    1 = First Song

		'Workaround for using "." as separator at discogs -----------------------------------------------------------------------------------------------------------
		tmp = 0 : tmp2 = 0
		NoSubTrackUsing = False
		For Each t In CurrentRelease("tracklist")
			Set currentTrack = CurrentRelease("tracklist")(t)
			position = currentTrack("position")
			If position <> "" Then
				tmp2 = tmp2 + 1
			End If
			If InStr(position, ".") <> 0 Then tmp = tmp + 1
		Next
		If tmp = tmp2 Then NoSubTrackUsing = True	'all tracks have "." in position tag, this can't be a subtrack
		'Workaround for using "." as separator at discogs -----------------------------------------------------------------------------------------------------------

		For Each t In CurrentRelease("tracklist")
			Set currentTrack = CurrentRelease("tracklist")(t)
			
			position = currentTrack("position")
			trackName = PackSpaces(DecodeHtmlChars(currentTrack("title")))
			Durations.Add currentTrack("duration")
			position = exchange_roman_numbers(position)
			
			' Here comes the new track/disc numbering methods
			If CheckUnselectNoTrackPos And position = "" Then
				UnselectedTracks(iTrackNum) = "x"
			End If
			
			If position <> "" Then
				If (cSubTrack <> -1 And InStr(LCase(position), ".") = 0 And CharSeparatorSubTrack = 1) Or (cSubTrack <> -1 And IsNumeric(Right(position, 1)) And CharSeparatorSubTrack = 2) Then
					If SubTrackNameSelection = False Then
						Tracks.Item(cSubTrack) = Tracks.Item(cSubTrack) & " (" & subTrackTitle & ")"
					Else
						Tracks.Item(cSubTrack) = subTrackTitle
					End If
					cSubTrack = -1
					subTrackTitle = ""
					CharSeparatorSubTrack = 0
				End If
				pos = 0
				If InStr(LCase(position), "-") > 0 Then
					pos = InStr(LCase(position), "-")
				End If
				'SubTrack Function ---------------------------------------------------------
				If InStr(LCase(position), ".") > 0 Then
					CharSeparatorSubTrack = 1
				End If
				If Not IsNumeric(Right(position, 1)) And Len(position) > 1 Then
					CharSeparatorSubTrack = 2
				End If
				If CharSeparatorSubTrack <> 0 Then
					If cSubTrack = -1 Then 'new subtrack
						If SubTrackNameSelection = False Then
							cSubTrack = iTrackNum - 1
						Else
							cSubTrack = iTrackNum
						End If
					End If
					If subTrackTitle = "" Then
						subTrackTitle = trackName
						UnselectedTracks(iTrackNum) = ""
					Else
						subTrackTitle = subTrackTitle & ", " & trackName
						UnselectedTracks(iTrackNum) = "x"
					End If
					If UserChoose = True Then
						UnselectedTracks(iTrackNum) = ""
					End If
					'SubTrack Function ---------------------------------------------------------
				End If
				If pos > 0 And CheckNoDisc = False Then ' Disc Number Included
					If CheckForceNumeric Then
						If Left(position,2) = "CD" Then
							If Mid(position,3,1) = "-" Then
								iAutoDiscNumber = 1
							Else
								If iAutoDiscNumber <> Mid(position,3,1) Then
									iAutoTrackNumber = 1
								End If
							End If
						End If
						If Left(position,2) <> "CD" And iAutoDiscNumber <> Left(position,pos-1) Then
							iAutoTrackNumber = 1
						End If
						If UnselectedTracks(iTrackNum) <> "x" Then
							If CheckLeadingZero = True And iAutoTrackNumber < 10 Then
								tracksNum.Add "0" & iAutoTrackNumber
							Else
								tracksNum.Add iAutoTrackNumber
							End If
							iAutoTrackNumber = iAutoTrackNumber + 1
						Else
							tracksNum.Add ""
						End If
					Else
						If pos > 0 Then
							If Len(Mid(position, pos+1)) > 1 Then	'minimum 2 Char after -  (1-1a, 1-II, 1-12)
								If IsInteger(Mid(position, pos+1, 1)) And Not IsInteger(Right(position, 1)) Then	'First is a Number, Char at the end (1-1a, 1-1b, 1-1c,...) = Sub-Track !
									If Mid(position,pos + 1, Len(position) - pos - 1) < 10 And CheckLeadingZero = True Then
										tracksNum.Add "0" & Right(position,Len(position)-pos)
									Else
										tracksNum.Add Right(position,Len(position)-pos)
									End If
								ElseIf IsInteger(Mid(position, pos+1)) Then		'no char at all (1-01, 1-02, 1-12)
									If CheckLeadingZero = True And Right(position,Len(position)-pos) < 10 Then
										tracksNum.Add "0" & Right(position,Len(position)-pos)
									Else
										tracksNum.Add Right(position,Len(position)-pos)
									End If
								Else
									tracksNum.Add Right(position,Len(position)-pos)
								End If
							ElseIf Len(Mid(position, pos+1)) = 1 Then	'1 Char after -  (1-1, 1-I, 1-2)
								If IsInteger(Mid(position, pos+1)) Then
									If CheckLeadingZero = True And Mid(position, pos+1) < 10 Then
										tracksNum.Add "0" & Mid(position, pos+1)
									Else
										tracksNum.Add Mid(position, pos+1)
									End If
								Else
									tracksNum.Add Mid(position, pos+1)
								End If
							End If
						End If
						If UnselectedTracks(iTrackNum) <> "x" Then
							If IsInteger(Right(position,Len(position)-pos)) Then
								iAutoTrackNumber = Right(position,Len(position)-pos) + 1
							Else
								iAutoTrackNumber = iAutoTrackNumber + 1
							End If
						End If
					End If
					If Left(position,2) = "CD" Then
						If Mid(position,3,1) = "-" Then
							'Or Mid(position,3,1) = "."
							iAutoDiscNumber = 1
						Else
							iAutoDiscNumber = Mid(position,3,1)
						End If
					End If
					If Left(position,2) <> "CD" Then iAutoDiscNumber = Left(position,pos-1)
					tracksCD.Add iAutoDiscNumber
				Else ' Apply Track Numbering Schemes
					If Not CheckSidesToDisc Or IsInteger(Left(position,1)) Then
						If CheckForceNumeric Then
							If UnselectedTracks(iTrackNum) <> "x" Then
								If CheckLeadingZero = True And iAutoTrackNumber < 10 Then
									tracksNum.Add "0" & iAutoTrackNumber
								Else
									tracksNum.Add iAutoTrackNumber
								End If
								iAutoTrackNumber = iAutoTrackNumber + 1
							Else
								tracksNum.Add ""
							End If
						Else
							If CheckLeadingZero = True And IsInteger(position) Then
								If position < 10 Then
									tracksNum.Add "0" & position
								Else
									tracksNum.Add position
								End If
							Else
								tracksNum.Add position
							End If
							If UnselectedTracks(iTrackNum) <> "x" Then
								If IsInteger(position) Then
									iAutoTrackNumber = position + 1
								Else
									iAutoTrackNumber = iAutoTrackNumber + 1
								End If
							End If
						End If
						If CheckForceDisc Then
							tracksCD.Add iAutoDiscNumber
						Else
							tracksCD.Add ""
						End If
					Else
						If Len(position) = 1 Then ' Only side is specified
							If CheckLeadingZero = True Then
								tracksNum.Add "01"
							Else
								tracksNum.Add "1"
							End If
							If 	LastDisc <> position Then
								If 	LastDisc <> "" Then
									iAutoDiscNumber = iAutoDiscNumber + 1
								End If
								LastDisc = position
							End If
							If CheckForceNumeric Then
								tracksCD.Add iAutoDiscNumber
							Else
								tracksCD.Add position
							End If
						ElseIf Len(position) = 2 Then
							If IsInteger(Mid(position,2,1)) And Not IsInteger(Mid(position,1,1)) Then
								' First is Side Second is Track
								If CheckLeadingZero = True And Mid(position,2) < 10 Then
									tracksNum.Add "0" & Mid(position,2)
								Else
									tracksNum.Add Mid(position,2)
								End If
								If 	LastDisc <>  Left(position,1) Then
									If 	LastDisc <> "" Then
										iAutoDiscNumber = iAutoDiscNumber + 1
									End If
									LastDisc = Left(position,1)
								End If
								If CheckForceNumeric Then
									tracksCD.Add iAutoDiscNumber
								Else
									tracksCD.Add Left(position,1)
								End If
							Else ' Two byte side
								tracksNum.Add "1"
								If 	LastDisc <>  position Then
									If 	LastDisc <> "" Then
										iAutoDiscNumber = iAutoDiscNumber + 1
									End If
									LastDisc = position
								End If
								If CheckForceNumeric Then
									tracksCD.Add iAutoDiscNumber
								Else
									tracksCD.Add position
								End If
							End If
						Else ' More than 2 bytes
							If IsInteger(Mid(position,2)) And CheckNoDisc = False Then
								'First is Side Latter is Track
								tracksNum.Add Mid(position,2)
								If 	LastDisc <>  Left(position,1) Then
									If 	LastDisc <> "" Then
										iAutoDiscNumber = iAutoDiscNumber + 1
									End If
									LastDisc = Left(position,1)
								End If
								If CheckForceNumeric Then
									tracksCD.Add iAutoDiscNumber
								Else
									tracksCD.Add Left(position,1)
								End If
							ElseIf IsInteger(Mid(position,3)) And CheckNoDisc = False Then
								' Two Byte Side, Latter is Track
								tracksNum.Add Mid(position,3)
								If 	LastDisc <>  Left(position,2) Then
									If 	LastDisc <> "" Then
										iAutoDiscNumber = iAutoDiscNumber + 1
									End If
									LastDisc = Left(position,2)
								End If
								If CheckForceNumeric Then
									tracksCD.Add iAutoDiscNumber
								Else
									tracksCD.Add Left(position,2)
								End If
							Else ' More than two non numeric bytes!
								If CheckNoDisc = False Then
									tracksNum.Add position
									tracksCD.Add ""
								Else
									If CheckForceNumeric Then
										If UnselectedTracks(iTrackNum) <> "x" Then
											If CheckLeadingZero = True And iAutoTrackNumber < 10 Then
												tracksNum.Add "0" & iAutoTrackNumber
											Else
												tracksNum.Add iAutoTrackNumber
											End If
											iAutoTrackNumber = iAutoTrackNumber + 1
										Else
											tracksNum.Add ""
										End If
									Else
										If UnselectedTracks(iTrackNum) <> "x" Then
											If IsInteger(position) Then
												tracksNum.Add iAutoTrackNumber
												iAutoTrackNumber = position + 1
											Else
												tracksNum.Add iAutoTrackNumber
												iAutoTrackNumber = iAutoTrackNumber + 1
											End If
										End If
									End If
									tracksCD.Add ""
								End If
							End If
						End If
					End If
				End If
			ElseIf currentTrack("duration") = "" And currentTrack("title") = "-" Then
				tracksNum.Add ""
				tracksCD.Add ""
				UnselectedTracks(iTrackNum) = "x"
			Else ' Nothing specified
				If CheckForceNumeric And UnselectedTracks(iTrackNum) <> "x" Then
					If CheckLeadingZero = True And iAutoTrackNumber < 10 Then
						tracksNum.Add "0" & iAutoTrackNumber
					Else
						tracksNum.Add iAutoTrackNumber
					End If
					iAutoTrackNumber = iAutoTrackNumber + 1
				Else
					tracksNum.Add ""
				End If
				If CheckForceDisc Then
					tracksCD.Add iAutoDiscNumber
				Else
					tracksCD.Add ""
				End If
			End If
			
			Dim involvedArtist, involvedTemp, involvedRole
			Dim TrackInvolvedPeople, TrackComposers, TrackConductors, TrackProducers, TrackLyricists, TrackFeaturing
			ReDim Involved_R_T(0)
			Dim ret
			
			TrackInvolvedPeople = ""
			TrackComposers = ""
			TrackConductors = ""
			TrackProducers = ""
			TrackLyricists = ""
			TrackFeaturing = AlbumFeaturing
			
			If UBound(Involved_R) > 0 Then
				For tmp = 1 To UBound(Involved_R)
					ReDim Preserve Involved_R_T(tmp)
					Involved_R_T(tmp) = Involved_R(tmp)
				Next
			End If
			
			For tmp = 1 To UBound(TrackPos)
				If TrackPos(tmp) = position Then
					involvedRole = TrackRoles(tmp)
					involvedArtist = TrackArtist2(tmp)
					
					If LookForFeaturing(involvedRole) Then
						If InStr(TrackFeaturing, involvedArtist) = 0 Then
							If TrackFeaturing = "" Then
								If CheckFeaturingName Then
									TrackFeaturing = TxtFeaturingName & " " & involvedArtist
								Else
									TrackFeaturing = involvedRole & " " & involvedArtist
								End If
							Else
								TrackFeaturing = TrackFeaturing & Separator & involvedArtist
							End If
						End If
					Else
						Do
							ret = searchKeyword(LyricistKeywords, involvedRole, TrackLyricists, involvedArtist)
							If ret <> "" And ret <> "ALREADY_INSIDE_ROLE" Then
								TrackLyricists = ret
								Exit Do
							ElseIf ret = "ALREADY_INSIDE_ROLE" Then
								Exit Do
							End If
							ret = searchKeyword(ConductorKeywords, involvedRole, TrackConductors, involvedArtist)
							If ret <> "" And ret <> "ALREADY_INSIDE_ROLE" Then
								TrackConductors = ret
								Exit Do
							ElseIf ret = "ALREADY_INSIDE_ROLE" Then
								Exit Do
							End If
							ret = searchKeyword(ProducerKeywords, involvedRole, TrackProducers, involvedArtist)
							If ret <> "" And ret <> "ALREADY_INSIDE_ROLE" Then
								TrackProducers = ret
								Exit Do
							ElseIf ret = "ALREADY_INSIDE_ROLE" Then
								Exit Do
							End If
							ret = searchKeyword(ComposerKeywords, involvedRole, TrackComposers, involvedArtist)
							If ret <> "" And ret <> "ALREADY_INSIDE_ROLE" Then
								TrackComposers = ret
								Exit Do
							ElseIf ret = "ALREADY_INSIDE_ROLE" Then
								Exit Do
							End If
							tmp2 = search_involved(Involved_R_T, involvedRole)
							If tmp2 = -1 Then
								ReDim Preserve Involved_R_T(UBound(Involved_R_T)+1)
								Involved_R_T(UBound(Involved_R_T)) = involvedRole & ": " & TrackArtist2(tmp)
							Else
								If InStr(Involved_R_T(tmp2), TrackArtist2(tmp)) = 0 Then
									Involved_R_T(tmp2) = Involved_R_T(tmp2) & ", " & TrackArtist2(tmp)
								End If
							End If
							Exit Do
						Loop While True
					End If
				End If
			Next
			
			Dim trackArtist, artistList, FoundFeaturing, tmpJoin, tmpTrackArtist
			artistList = ""
			tmpJoin = ""
			
			If currentTrack.Exists("artists") Then
				FoundFeaturing = False
				For Each artist In currentTrack("artists")
					Set currentArtist = currentTrack("artists")(artist)
					If (currentArtist("anv") <> "") And Not CheckUseAnv Then
						tmpTrackArtist = CleanArtistName(currentArtist("anv"))
					Else
						tmpTrackArtist = CleanArtistName(currentArtist("name"))
					End If
					If FoundFeaturing = False Then
						artistList = artistList & tmpTrackArtist
					Else
						If TrackFeaturing = "" Then
							If CheckFeaturingName Then
								TrackFeaturing = TxtFeaturingName & " " & tmpTrackArtist
							Else
								TrackFeaturing = tmpJoin & " " & tmpTrackArtist
							End If
						Else
							TrackFeaturing = TrackFeaturing & ", " & tmpTrackArtist
						End If

					End If
					'TitleFeaturing
					If currentArtist("join") <> "" Then
						If LookForFeaturing(currentArtist("join")) Then
							FoundFeaturing = True
							tmpJoin = currentArtist("join")
						Else
							artistList = artistList & " " & currentArtist("join") & " "
							FoundFeaturing = False
						End If
					End If
				Next
			End If
			If Right(artistList, 3) = " , " Then artistList = Left(artistList, Len(artistList)-3)

			If artistList = "" Then artistList = AlbumArtistTitle

			Dim extra
			If currentTrack.Exists("extraartists") Then
				For Each extra In currentTrack("extraartists")
					Set currentArtist = CurrentTrack("extraartists")(extra)
					If (currentArtist("anv") <> "") And Not CheckUseAnv Then
						involvedArtist = CleanArtistName(currentArtist("anv"))
					Else
						involvedArtist = CleanArtistName(currentArtist("name"))
					End If
					If involvedArtist <> "" Then
						role = currentArtist("role")
						NoSplit = False
						If InStr(role, ",") = 0 Then
							involvedRole = role
							zahl = 0
							NoSplit = True
						Else
							rolea = Split(role, ", ")
							zahl = UBound(rolea)
						End If
						For zahltemp = 0 To zahl
							If NoSplit = False Then
								involvedRole = rolea(zahltemp)
							End If
							
							If LookForFeaturing(involvedRole) Then
								If InStr(artistList, involvedArtist) = 0 Then
									If TrackFeaturing = "" Then
										If CheckFeaturingName Then
											TrackFeaturing = TxtFeaturingName & " " & involvedArtist
										Else
											TrackFeaturing = involvedRole & " " & involvedArtist
										End If
									Else
										If InStr(TrackFeaturing, involvedArtist) = 0 Then
											TrackFeaturing = TrackFeaturing & ", " & involvedArtist
										End If
									End If
								End If
							Else
								Do
									tmp = searchKeyword(LyricistKeywords, involvedRole, TrackLyricists, involvedArtist)
									If tmp <> "" And tmp <> "ALREADY_INSIDE_ROLE" Then
										TrackLyricists = tmp
										Exit Do
									ElseIf tmp = "ALREADY_INSIDE_ROLE" Then
										Exit Do
									End If
									tmp = searchKeyword(ConductorKeywords, involvedRole, TrackConductors, involvedArtist)
									If tmp <> "" And tmp <> "ALREADY_INSIDE_ROLE" Then
										TrackConductors = tmp
										Exit Do
									ElseIf tmp = "ALREADY_INSIDE_ROLE" Then
										Exit Do
									End If
									tmp = searchKeyword(ProducerKeywords, involvedRole, TrackProducers, involvedArtist)
									If tmp <> "" And tmp <> "ALREADY_INSIDE_ROLE" Then
										TrackProducers = tmp
										Exit Do
									ElseIf tmp = "ALREADY_INSIDE_ROLE" Then
										Exit Do
									End If
									tmp = searchKeyword(ComposerKeywords, involvedRole, TrackComposers, involvedArtist)
									If tmp <> "" And tmp <> "ALREADY_INSIDE_ROLE" Then
										TrackComposers = tmp
										Exit Do
									ElseIf tmp = "ALREADY_INSIDE_ROLE" Then
										Exit Do
									End If
									tmp2 = search_involved(Involved_R_T, involvedRole)
									If tmp2 = -1 Then
										ReDim Preserve Involved_R_T(UBound(Involved_R_T)+1)
										Involved_R_T(UBound(Involved_R_T)) = involvedRole & ": " & involvedArtist
									Else
										If InStr(Involved_R_T(tmp2), involvedArtist) = 0 Then
											Involved_R_T(tmp2) = Involved_R_T(tmp2) & ", " & involvedArtist
										End If
									End If
									Exit Do
								Loop While True
							End If
						Next
					End If
				Next
			End If
			
			If TrackFeaturing <> "" Then
				If CheckTitleFeaturing = True Then
					tmp = InStrRev(TrackFeaturing, ", ")
					If tmp = 0 Then
						trackName = trackName & " (" & TrackFeaturing & ")"
					Else
						trackName = trackName & " (" &  Left(TrackFeaturing, tmp-1) & " & " & Mid(TrackFeaturing, tmp+2) & ")"
					End If
				Else
					tmp = InStrRev(TrackFeaturing, ", ")
					If tmp = 0 Then
						artistList = artistList & " " & TrackFeaturing
					Else
						artistList = artistList & " " & Left(TrackFeaturing, tmp-1) & " & " & Mid(TrackFeaturing, tmp+2)
					End If
				End If
			End If
			
			If ArtistSeparator <> ", " Then
				artistList = Replace(artistList, ", ", ArtistSeparator)
				artistList = Replace(artistList, " " & ArtistSeparator, ArtistSeparator)
			Else
				artistList = Replace(artistList, " , ", ", ")
			End If
			ArtistTitles.Add artistList
			
			TrackLyricists = FindArtist(TrackLyricists, AlbumLyricist)
			If AlbumLyricist <> "" And TrackLyricists <> "" Then
				Lyricists.Add AlbumLyricist & "; " & TrackLyricists
			Else
				Lyricists.Add AlbumLyricist & TrackLyricists
			End If
			TrackComposers = FindArtist(TrackComposers, AlbumComposer)
			If AlbumComposer <> "" And TrackComposers <> "" Then
				Composers.Add AlbumComposer & "; " & TrackComposers
			Else
				Composers.Add AlbumComposer & TrackComposers
			End If
			TrackConductors = FindArtist(TrackConductors, AlbumConductor)
			If AlbumConductor <> "" And TrackConductors <> "" Then
				Conductors.Add AlbumConductor & "; " & TrackConductors
			Else
				Conductors.Add AlbumConductor & TrackConductors
			End If
			TrackProducers = FindArtist(TrackProducers, AlbumProducer)
			If AlbumProducer <> "" And TrackProducers <> "" Then
				Producers.Add AlbumProducer & "; " & TrackProducers
			Else
				Producers.Add AlbumProducer & TrackProducers
			End If
			
			If UBound(Involved_R_T) > 0 Then
				For tmp = 1 To UBound(involved_R_T)
					TrackInvolvedPeople = TrackInvolvedPeople & Involved_R_T(tmp) & "; "
				Next
				TrackInvolvedPeople = Left(TrackInvolvedPeople, Len(TrackInvolvedPeople)-2)
			Else
				TrackInvolvedPeople = ""
			End If
			
			InvolvedArtists.Add TrackInvolvedPeople
			Tracks.Add trackName
			iTrackNum = iTrackNum + 1
		Next
		
		If cSubTrack <> -1 Then
			If SubTrackNameSelection = False Then
				Tracks.Item(cSubTrack) = Tracks.Item(cSubTrack) & " (" & subTrackTitle & ")"
			Else
				Tracks.Item(cSubTrack) = subTrackTitle
			End If
			cSubTrack = -1
			subTrackTitle = ""
			CharSeparatorSubTrack = 0
		End If
		
		' Get album title
		AlbumTitle = currentRelease("title")
		
		' Get Album art URL
		AlbumArtThumbnail = CurrentRelease("thumb")
		AlbumArtThumbnail = Replace(AlbumArtThumbnail, "http://api.discogs.com", "http://s.pixogs.com")
		
		Dim i, currentImage
		If CurrentRelease.Exists("images") Then
			For Each i In CurrentRelease("images")
				Set currentImage = CurrentRelease("images")(i)
				
				If currentImage("type") = "primary" Or AlbumArtURL = "" Then
					AlbumArtURL = currentImage("uri")
					AlbumArtURL = Replace(AlbumArtURL, "http://api.discogs.com", "http://s.pixogs.com")
				End If
			Next
		End If
		
		'----------------------------------DiscogsImages----------------------------------------
		Rem Set ImageList = SDB.NewStringList
		Rem Set SaveImageType = SDB.NewStringList
		Rem Set SaveImage = SDB.NewStringList
		Rem ImagesCount = 0
		Rem Dim FirstAlbumArt, tmpArt
		
		Rem If CurrentRelease.Exists("images") Then
		Rem ImagesCount = CurrentRelease("images").Count
		Rem If CurrentRelease("images").Count > 1 Then
		Rem For Each i In CurrentRelease("images")
		Rem Set currentImage = CurrentRelease("images")(i)
		Rem If currentImage("type") = "primary" Then
		Rem FirstAlbumArt = currentImage("uri")
		Rem FirstAlbumArt = Replace(FirstAlbumArt, "http://api.discogs.com", "http://s.pixogs.com")
		Rem End If
		Rem Next
		Rem If FirstAlbumArt = "" Then
		Rem FirstAlbumArt = currentImage("uri")
		Rem FirstAlbumArt = Replace(FirstAlbumArt, "http://api.discogs.com", "http://s.pixogs.com")
		Rem End If
		Rem For Each i In CurrentRelease("images")
		Rem Set currentImage = CurrentRelease("images")(i)
		Rem tmpArt = currentImage("uri")
		Rem tmpArt = Replace(tmpArt, "http://api.discogs.com", "http://s.pixogs.com")
		Rem If FirstAlbumArt <> tmpArt Then
		Rem ImageList.add tmpArt
		Rem SaveImageType.add "other"
		Rem SaveImage.add "0"
		Rem End If
		Rem Next
		Rem End If
		Rem End If
		'----------------------------------DiscogsImages----------------------------------------
		
		' Get Master ID
		Dim theMaster
		If CurrentRelease.Exists("master_id") Then
			theMaster = currentRelease("master_id")
			If SavedMasterId <> theMaster Then
				OriginalDate = ReloadMaster(theMaster)
				SavedMasterId = theMaster
			End If
		Else
			theMaster = ""
			SavedMasterId = theMaster
			OriginalDate = ""
		End If
		
		Dim ReleaseSplit
		' Get release year/date
		If CurrentRelease.Exists("released") Then
			ReleaseDate = CurrentRelease("released")
			If Len(ReleaseDate) > 4 Then
				ReleaseSplit = Split(ReleaseDate,"-")
				If ReleaseSplit(2) = "00" Then
					ReleaseDate = Left(ReleaseDate, 4)
				Else
					ReleaseDate = ReleaseSplit(2) & "-" & ReleaseSplit(1) & "-" & ReleaseSplit(0)
				End If
				If CheckYearOnlyDate Then
					ReleaseDate = Right(ReleaseDate, 4)
				End If
			End If
		Else
			ReleaseDate = ""
		End If
		
		'Set OriginalDate
		If OriginalDate <> "" Then
			If Len(OriginalDate) > 4 Then
				ReleaseSplit = Split(OriginalDate,"-")
				If ReleaseSplit(2) = "00" Then
					OriginalDate = Left(OriginalDate, 4)
				Else
					OriginalDate = ReleaseSplit(2) & "-" & ReleaseSplit(1) & "-" & ReleaseSplit(0)
				End If
				If CheckYearOnlyDate Then
					OriginalDate = Right(OriginalDate, 4)
				End If
			End If
		End If
		
		' Get genres
		Dim g
		For Each g In CurrentRelease("genres")
			AddToField Genres, CurrentRelease("genres")(g)
		Next
		
		' Get styles/moods/themes
		Dim s
		If CurrentRelease.Exists("styles") Then
			For Each s In CurrentRelease("styles")
				AddToField Styles, CurrentRelease("styles")(s)
			Next
		End If
		
		' Get Label
		Dim l, currentLabel
		If CurrentRelease.Exists("labels") Then
			For Each l In CurrentRelease("labels")
				Set currentLabel = CurrentRelease("labels")(l)
				If SavedLabelId = "" Then
					If currentLabel.Exists("id") Then
						SavedLabelId = currentLabel("id")
					End If
				End If
				AddToField theLabels, CleanArtistName(currentLabel("name"))
				AddToField theCatalogs, currentLabel("catno")
			Next
		Else
			theLabels = ""
			theCatalogs = ""
		End If
		
		' Get Country
		If CurrentRelease.Exists("country") Then
			theCountry = CurrentRelease("country")
		Else
			theCountry = ""
		End If
		
		' Get Format
		Dim f, d, currentFormat
		If CurrentRelease.Exists("formats") Then
			For Each f In CurrentRelease("formats")
				Set currentFormat = CurrentRelease("formats")(f)
				AddToField theFormat, currentFormat("name")
				If currentFormat.Exists("descriptions") Then
					For Each d In currentFormat("descriptions")
						theFormat = theFormat & ", " & currentFormat("descriptions")(d)
					Next
				End If
			Next
		Else
			theFormat = ""
		End If
		
		' Get Comment
		If CurrentRelease.Exists("notes") Then
			Comment = CurrentRelease("notes")
		Else
			Comment = ""
		End If
		
		' Get data_quality
		Dim DataQuality
		If CurrentRelease.Exists("data_quality") Then
			DataQuality = CurrentRelease("data_quality")
		Else
			DataQuality = ""
		End If
	End If
	
	
	FormatSearchResultsViewer Tracks, TracksNum, TracksCD, Durations, AlbumArtist, AlbumArtistTitle, ArtistTitles, AlbumTitle, ReleaseDate, OriginalDate, Genres, Styles, theLabels, theCountry, AlbumArtThumbNail, CurrentReleaseID, theCatalogs, Lyricists, Composers, Conductors, Producers, InvolvedArtists, theFormat, theMaster, comment, DiscogsTracksNum, DataQuality
	
	ret = trackliste_aufbauen(Tracks, TracksNum, TracksCD, Durations, AlbumArtist, AlbumArtistTitle, ArtistTitles, AlbumTitle, ReleaseDate, OriginalDate, Genres, Styles, theLabels, theCountry, AlbumArtThumbNail, CurrentReleaseID, theCatalogs, Lyricists, Composers, Conductors, Producers, InvolvedArtists, theFormat, theMaster, Comment)
	If ret = True And SkipNotChangedReleases = True Then
		WriteLog "SKIP ReloadResults !!!"
		CurrentSelectedAlbum = CurrentSelectedAlbum + 1
		cReleasesAutoSkip = cReleasesAutoSkip + 1
		WriteLog("AlbumIDList.Count=" & AlbumIDList.Count)
		If AlbumIDList.Count < CurrentSelectedAlbum Then
			FinishSearch Form
		Else
			NewSearch CurrentSelectedAlbum
		End If
	Else
		WebBrowser2.SetHTMLDocument ""
		WebBrowser2.SetHTMLDocument tracklistHTML
		WriteLog ("RadioBoxCheck=" & RadioBoxCheck)
		If RadioBoxCheck > -1 Then
			Dim RadioBox, templateHTMLDoc
			Set templateHTMLDoc = WebBrowser2.Interf.Document
			Set RadioBox = templateHTMLDoc.getElementById(RadioBoxCheck)
			RadioBox.Checked = True
		End If
		WriteLog "Stop ReloadResults"
	End If
	
End Sub


Function FindArtist(ArtistList1, ArtistList2)
	
	Dim tmpArtist, i
	ReDim newArtistList1(0)
	tmpArtist = Split(ArtistList1, "; ")
	For i = 0 To UBound(tmpArtist)
		If InStr(ArtistList2,tmpArtist(i)) = 0 Then
			ReDim Preserve newArtistList1(UBound(newArtistList1)+1)
			newArtistList1(UBound(newArtistList1)) = tmpArtist(i)
		End If
	Next
	For i = 0 To UBound(newArtistList1)
		If FindArtist = "" Then
			FindArtist = newArtistList1(i)
		Else
			FindArtist = FindArtist & "; " & newArtistList1(i)
		End If
	Next
	
End Function


Sub Track_from_to (currentTrack, currentArtist, involvedRole, Title_Position, TrackRoles, TrackArtist2, TrackPos, LeadingZeroTrackPosition)
	
	Dim tmp3, tmp4, tmpSide1, tmpSide2, tmpSideD, Vinyl_Pos1, Vinyl_Pos2, zahltemp3, ret
	tmp3 = Split(currentTrack, " ")
	tmpSide1 = ""
	tmpSide2 = ""
	tmpSideD = ""
	
	tmp3(0) = exchange_roman_numbers(tmp3(0))
	tmp3(2) = exchange_roman_numbers(tmp3(2))
	
	If InStr(tmp3(0), "-") <> 0 Then
		tmp4 = Split(tmp3(0), "-")
		tmpSide1 = tmp4(0)
		tmp3(0) = tmp4(1)
		tmp3(0) = exchange_roman_numbers(tmp3(0))
		tmpSideD = "-"
	End If
	If InStr(tmp3(2), "-") <> 0 Then
		tmp4 = Split(tmp3(2), "-")
		tmpSide2 = tmp4(0)
		tmp3(2) = tmp4(1)
		tmp3(2) = exchange_roman_numbers(tmp3(2))
		tmpSideD = "-"
	End If
	If InStr(tmp3(0), ".") <> 0 Then
		tmp4 = Split(tmp3(0), ".")
		tmpSide1 = tmp4(0)
		tmp3(0) = tmp4(1)
		tmp3(0) = exchange_roman_numbers(tmp3(0))
		tmpSideD = "."
	End If
	If InStr(tmp3(2), ".") <> 0 Then
		tmp4 = Split(tmp3(2), ".")
		tmpSide2 = tmp4(0)
		tmp3(2) = tmp4(1)
		tmp3(2) = exchange_roman_numbers(tmp3(2))
		tmpSideD = "."
	End If
	If Left(tmp3(0), 2) = "CD" Then
		tmpSide1 = "CD"
		tmp3(0) = Mid(tmp3(0), 3)
	End If
	If Left(tmp3(2), 2) = "CD" Then
		tmpSide2 = "CD"
		tmp3(2) = Mid(tmp3(2), 3)
	End If
	If Left(tmp3(0), 3) = "DVD" Then
		tmpSide1 = "DVD"
		tmp3(0) = Mid(tmp3(0), 4)
	End If
	If Left(tmp3(2), 3) = "DVD" Then
		tmpSide2 = "DVD"
		tmp3(2) = Mid(tmp3(2), 4)
	End If
	If IsNumeric(Right(tmp3(0),1)) = False And Len(tmp3(0)) > 1 Then
		tmp3(0) = Left(tmp3(0), Len(tmp3(0))-1)
	End If
	If IsNumeric(Right(tmp3(2),1)) = False And Len(tmp3(2)) > 1 Then
		tmp3(2) = Left(tmp3(2), Len(tmp3(2))-1)
	End If
	If IsNumeric(tmp3(0)) = False Then
		If Len(tmp3(0)) > 1 Then
			tmpSide1 = Left(tmp3(0), 1)
			tmp3(0) = Mid(tmp3(0), 2)
		Else
			tmpSide1 = tmp3(0)
			tmp3(0) = 1
		End If
	End If
	If IsNumeric(tmp3(2)) = False Then
		If Len(tmp3(2)) > 1 Then
			tmpSide2 = Left(tmp3(2), 1)
			tmp3(2) = Mid(tmp3(2), 2)
		Else
			tmpSide2 = tmp3(2)
			tmp3(2) = 1
		End If
	End If
	
	If tmpSide1 <> tmpSide2 Then
		Vinyl_Pos1 = tmpSide1
		Vinyl_Pos2 = tmp3(0)
		Do
			If LeadingZeroTrackPosition = True And Vinyl_Pos2 < 10 Then
				Vinyl_Pos2 = "0" & Vinyl_Pos2
			End If
			tmp4 = Vinyl_Pos1 & tmpSideD & Vinyl_Pos2
			ret = search_involved(Title_Position, tmp4)
			If ret = -1 Then
				If IsNumeric(Vinyl_Pos1) = True Then
					If Vinyl_Pos1 > 101 Then Exit Do
					Vinyl_Pos1 = Vinyl_Pos1 + 1
				Else
					If Chr(Asc(Vinyl_Pos1)) = "Z" Then Exit Do
					Vinyl_Pos1 = Chr(Asc(Vinyl_Pos1) + 1)
				End If
				Vinyl_Pos2 = "1"
			Else
				ReDim Preserve TrackRoles(UBound(TrackRoles)+1)
				ReDim Preserve TrackArtist2(UBound(TrackArtist2)+1)
				ReDim Preserve TrackPos(UBound(TrackPos)+1)
				TrackArtist2(UBound(TrackArtist2)) = currentArtist
				TrackRoles(UBound(TrackRoles)) = involvedRole
				TrackPos(UBound(TrackPos)) = Vinyl_Pos1 & tmpSideD & Vinyl_Pos2
				If CStr(Vinyl_Pos1) = CStr(tmpSide2) And CStr(Vinyl_Pos2) = CStr(tmp3(2)) Then Exit Do
				Vinyl_Pos2 = Vinyl_Pos2 + 1
			End If
		Loop While True
	Else
		For zahltemp3 = tmp3(0) To tmp3(2)
			If LeadingZeroTrackPosition = True And zahltemp3 < 10 Then
				zahltemp3 = "0" & zahltemp3
			End If
			ReDim Preserve TrackRoles(UBound(TrackRoles)+1)
			ReDim Preserve TrackArtist2(UBound(TrackArtist2)+1)
			ReDim Preserve TrackPos(UBound(TrackPos)+1)
			TrackArtist2(UBound(TrackArtist2)) = currentArtist
			TrackRoles(UBound(TrackRoles)) = involvedRole
			TrackPos(UBound(TrackPos)) = tmpSide1 & tmpSideD & zahltemp3
		Next
	End If
	
End Sub


Sub Add_Track_Role(currentTrack, currentArtist, involvedRole, TrackRoles, TrackArtist2, TrackPos)
	
	WriteLog "currentTrack=" & currentTrack
	currentTrack = exchange_roman_numbers(currentTrack)
	ReDim Preserve TrackRoles(UBound(TrackRoles)+1)
	ReDim Preserve TrackArtist2(UBound(TrackArtist2)+1)
	ReDim Preserve TrackPos(UBound(TrackPos)+1)
	TrackArtist2(UBound(TrackArtist2)) = currentArtist
	TrackRoles(UBound(TrackRoles)) = involvedRole
	TrackPos(UBound(TrackPos)) = currentTrack
	
End Sub


' ShowResult is called every time the search result is changed from the drop
' down at the top of the window
Sub ShowResult(ResultID)
	
	WriteLog "Start ShowResult"
	Dim searchURL, ReleaseID
	If ResultsReleaseID.Count = 0 Then
		Rem FormatErrorMessage "Cannot load artist - "
		WriteLog "Stop ShowResult"
		Exit Sub
	End If
	WebBrowser.SetHTMLDocument ""                 ' Deletes visible search result
	WebBrowser2.SetHTMLDocument ""
	
	ReleaseID = ResultsReleaseID.Item(ResultID)
	WriteLog "ShowResult ReleaseID=" & ReleaseID

	' use json api with vbsjson class at start of file now
	Dim oXMLHTTP : Set oXMLHTTP = CreateObject("MSXML2.XMLHTTP.6.0")
	
	Dim json
	Set json = New VbsJson
	
	Dim response
	
	searchURL = "http://api.discogs.com/releases/" & ReleaseID
	
	Call oXMLHTTP.open("GET", searchURL, False)
	Call oXMLHTTP.setRequestHeader("Content-Type","application/json")
	Call oXMLHTTP.setRequestHeader("User-Agent","MediaMonkeyDiscogsAutoTagBatch/2.0 +http://mediamonkey.com")
	Call oXMLHTTP.send()
	
	If oXMLHTTP.Status = 200 Then
		Set CurrentRelease = json.Decode(oXMLHTTP.responseText)
		
		CurrentReleaseID = ReleaseID
		UserChoose = False
		
		ReloadResults
	End If
	WriteLog "Stop ShowResult"
	
End Sub


' This does the final clean up, so that our script doesn't leave any unwanted traces
Sub FinishSearch(Panel)
	
	WriteLog "FinishSearch"

	If IsObject(WebBrowser) Then
		WebBrowser.Common.DestroyControl      ' Destroy the external control
	End If
	Set WebBrowser = Nothing              ' Release global variable
	SDB.Objects("WebBrowser") = Nothing
	If IsObject(WebBrowser2) Then
		WebBrowser2.Common.DestroyControl      ' Destroy the external control
	End If
	Set WebBrowser2 = Nothing              ' Release global variable
	If IsObject(SDB.Objects("WebBrowser2")) Then SDB.Objects("WebBrowser2") = Nothing
	SDB.Objects("SearchForm") = Nothing


	SDB.Objects("WebBrowser3") = Nothing
	Dim res, mtInformation
	res = SDB.MessageBox("You selected " & AlbumIDList.Count & " Albums. " & vbCr & cReleasesUpdate & " Albums were updated." & vbCr &  cReleasesSkip & " Albums were manually skipped." & vbCr & cReleasesAutoSkip & " Albums were auto skipped" & vbCr & cReleasesOnlyDiscogsSkip & " Albums were no Discogs-Releases" & vbCr & cReleasesNoDiscogsSkip & " Albums were Discogs Releases", mtInformation, Array(mbOk))
	Set ini = Nothing
	Set ResultsReleaseID = Nothing
	Script.UnregisterAllEvents

End Sub


Function GetHeader()
	
	Dim templateHTML2, i
	templateHTML2 = "<HTML>"
	templateHTML2 = templateHTML2 &  "<HEAD>"
	templateHTML2 = templateHTML2 &  "<style type=""text/css"" media=""screen"">"
	templateHTML2 = templateHTML2 &  ".tabletext { font-family: Arial, Helvetica, sans-serif; font-size: 8pt;}"
	templateHTML2 = templateHTML2 &  "</style>"
	templateHTML2 = templateHTML2 &  "</HEAD>"
	templateHTML2 = templateHTML2 &  "<body bgcolor=""#FFFFFF"">"
	templateHTML2 = templateHTML2 &  "<table border=0 width=100% cellspacing=0 cellpadding=1 class=tabletext>"
	templateHTML2 = templateHTML2 &  "<tr>"
	templateHTML2 = templateHTML2 &  "<td align=left><a href=""http://www.discogs.com"" target=""_blank""><img src=""http://s.discogss.com/images/discogs-white-2.png"" border=""0""/ alt=""Discogs Homepage""></a><b>" & VersionStr & "</b></td>"
	templateHTML2 = templateHTML2 &  "<td colspan=3 align=right valign=top>"
	
	templateHTML2 = templateHTML2 &  "<table border=0 cellspacing=0 cellpadding=2 class=tabletext>"
	templateHTML2 = templateHTML2 &  "<tr><td colspan=2></td><td><b>Filter Results: </b></td><td colspan=3> </td></tr>"
	templateHTML2 = templateHTML2 &  "<tr>"
	templateHTML2 = templateHTML2 &  "<td><b>Load:</b></td>"
	templateHTML2 = templateHTML2 &  "<td><b>Quick Search:</b></td>"
	templateHTML2 = templateHTML2 &  "<td align=left><button type=button class=tabletext id=""showmediatypefilter"">Set Type Filter</button></td>"
	templateHTML2 = templateHTML2 &  "<td align=left><button type=button class=tabletext id=""showmediaformatfilter"">Set Format Filter</button></td>"
	templateHTML2 = templateHTML2 &  "<td align=left><button type=button class=tabletext id=""showcountryfilter"">Set Country Filter</button></td>"
	templateHTML2 = templateHTML2 &  "<td align=left><button type=button class=tabletext id=""showyearfilter"">Set Year Filter</button></td>"
	templateHTML2 = templateHTML2 &  "</tr>"
	templateHTML2 = templateHTML2 &  "<tr>"
	templateHTML2 = templateHTML2 &  "<td>"
	templateHTML2 = templateHTML2 &  "<select id=""load"" class=tabletext title=""Search Result=Search with Artist and Album Title" & vbCrLf & "Master Release=Show all releases from the master"">"
	
	For i = 0 To LoadList.Count - 1
		If LoadList.Item(i) <> CurrentLoadType Then
			templateHTML2 = templateHTML2 &  "<option value=""" & EncodeHtmlChars(LoadList.Item(i)) & """>" & LoadList.Item(i) & "</option>"
		Else
			templateHTML2 = templateHTML2 &  "<option value=""" & EncodeHtmlChars(LoadList.Item(i)) & """ selected>" & LoadList.Item(i) & "</option>"
		End If
	Next
	templateHTML2 = templateHTML2 &  "</select>"
	templateHTML2 = templateHTML2 &  "</td>"
	'Alternative Searches Begin
	templateHTML2 = templateHTML2 &  "<td>"
	templateHTML2 = templateHTML2 &  "<select id=""alternative"" class=tabletext>"
	For i = 0 To AlternativeList.Count - 1
		If AlternativeList.Item(i) <> SavedSearchTerm Then
			templateHTML2 = templateHTML2 &  "<option value=""" & EncodeHtmlChars(AlternativeList.Item(i)) & """>" & AlternativeList.Item(i) & "</option>"
		Else
			templateHTML2 = templateHTML2 &  "<option value=""" & EncodeHtmlChars(AlternativeList.Item(i)) & """ selected>" & AlternativeList.Item(i) & "</option>"
		End If
	Next
	templateHTML2 = templateHTML2 &  "</select>"
	templateHTML2 = templateHTML2 &  "</td>"
	'Alternative Searches End
	'Filters Begin
	templateHTML2 = templateHTML2 &  "<td>"
	templateHTML2 = templateHTML2 &  "<select id=""filtermediatype"" class=tabletext>"
	
	If FilterMediaType = "None" Then
		templateHTML2 = templateHTML2 &  "<option value=""None"">No MediaType Filter</option>"
		templateHTML2 = templateHTML2 &  "<option style=""background-color:#F4113F;"" value=""Use MediaType Filter"">Use MediaType Filter</option>"
	ElseIf FilterMediaType = "Use MediaType Filter" Then
		templateHTML2 = templateHTML2 &  "<option style=""background-color:#F4113F;"" value=""Use MediaType Filter"">Use MediaType Filter</option>"
		templateHTML2 = templateHTML2 &  "<option value=""None"">No MediaType Filter</option>"
	End If
	If FilterMediaType <> "None" And FilterMediaType <> "Use MediaType Filter" Then
		templateHTML2 = templateHTML2 &  "<option value=""None"">No MediaType Filter</option>"
		templateHTML2 = templateHTML2 &  "<option value=""Use MediaType Filter"">Use MediaType Filter</option>"
	End If
	For i = 0 To MediaTypeList.Count - 1
		If FilterMediaType <> MediaTypeList.Item(i) Or FilterMediaType = "None" Or FilterMediaType = "Use MediaType Filter" Then
			templateHTML2 = templateHTML2 &  "<option value=""" & EncodeHtmlChars(MediaTypeList.Item(i)) & """>" & MediaTypeList.Item(i) & "</option>"
		Else
			templateHTML2 = templateHTML2 &  "<option value=""" & EncodeHtmlChars(MediaTypeList.Item(i)) & """ selected>" & MediaTypeList.Item(i) & "</option>"
		End If
	Next
	templateHTML2 = templateHTML2 &  "</select>"
	templateHTML2 = templateHTML2 &  "</td>"
	
	templateHTML2 = templateHTML2 &  "<td>"
	templateHTML2 = templateHTML2 &  "<select id=""filtermediaformat"" class=tabletext>"
	
	If FilterMediaFormat = "None" Then
		templateHTML2 = templateHTML2 &  "<option value=""None"">No MediaFormat Filter</option>"
		templateHTML2 = templateHTML2 &  "<option style=""background-color:#F4113F;"" value=""Use MediaFormat Filter"">Use MediaFormat Filter</option>"
	ElseIf FilterMediaFormat = "Use MediaFormat Filter" Then
		templateHTML2 = templateHTML2 &  "<option style=""background-color:#F4113F;"" value=""Use MediaFormat Filter"">Use MediaFormat Filter</option>"
		templateHTML2 = templateHTML2 &  "<option value=""None"">No MediaFormat Filter</option>"
	End If
	If FilterMediaFormat <> "None" And FilterMediaFormat <> "Use MediaFormat Filter" Then
		templateHTML2 = templateHTML2 &  "<option value=""None"">No MediaFormat Filter</option>"
		templateHTML2 = templateHTML2 &  "<option value=""Use MediaFormat Filter"">Use MediaFormat Filter</option>"
	End If
	For i = 1 To MediaFormatList.Count - 1
		If FilterMediaFormat <> MediaFormatList.Item(i) Or FilterMediaFormat = "None" Or FilterMediaFormat = "Use MediaFormat Filter" Then
			templateHTML2 = templateHTML2 &  "<option value=""" & EncodeHtmlChars(MediaFormatList.Item(i)) & """>" & MediaFormatList.Item(i) & "</option>"
		Else
			templateHTML2 = templateHTML2 &  "<option value=""" & EncodeHtmlChars(MediaFormatList.Item(i)) & """ selected>" & MediaFormatList.Item(i) & "</option>"
		End If
	Next
	
	templateHTML2 = templateHTML2 &  "</select>"
	templateHTML2 = templateHTML2 &  "</td>"
	
	templateHTML2 = templateHTML2 &  "<td>"
	templateHTML2 = templateHTML2 &  "<select id=""filtercountry"" class=tabletext>"
	
	If FilterCountry = "None" Then
		templateHTML2 = templateHTML2 &  "<option value=""None"">No Country Filter</option>"
		templateHTML2 = templateHTML2 &  "<option style=""background-color:#F4113F;"" value=""Use Country Filter"">Use Country Filter</option>"
	ElseIf FilterCountry = "Use Country Filter" Then
		templateHTML2 = templateHTML2 &  "<option style=""background-color:#F4113F;"" value=""Use Country Filter"">Use Country Filter</option>"
		templateHTML2 = templateHTML2 &  "<option value=""None"">No Country Filter</option>"
	End If
	If FilterCountry <> "None" And FilterCountry <> "Use Country Filter" Then
		templateHTML2 = templateHTML2 &  "<option value=""None"">No Country Filter</option>"
		templateHTML2 = templateHTML2 &  "<option value=""Use Country Filter"">Use Country Filter</option>"
	End If
	For i = 1 To CountryList.Count - 1
		If FilterCountry <> CountryList.Item(i) Or FilterCountry = "None" Or FilterCountry = "Use Country Filter" Then
			templateHTML2 = templateHTML2 &  "<option value=""" & EncodeHtmlChars(CountryList.Item(i)) & """>" & CountryList.Item(i) & "</option>"
		Else
			templateHTML2 = templateHTML2 &  "<option value=""" & EncodeHtmlChars(CountryList.Item(i)) & """ selected>" & CountryList.Item(i) & "</option>"
		End If
	Next
	
	templateHTML2 = templateHTML2 &  "</select>"
	templateHTML2 = templateHTML2 &  "</td>"
	
	templateHTML2 = templateHTML2 &  "<td>"
	templateHTML2 = templateHTML2 &  "<select id=""filteryear"" class=tabletext>"
	
	If FilterYear = "None" Then
		templateHTML2 = templateHTML2 &  "<option value=""None"">No Year Filter</option>"
		templateHTML2 = templateHTML2 &  "<option style=""background-color:#F4113F;"" value=""Use Year Filter"">Use Year Filter</option>"
	ElseIf FilterYear = "Use Year Filter" Then
		templateHTML2 = templateHTML2 &  "<option style=""background-color:#F4113F;"" value=""Use Year Filter"">Use Year Filter</option>"
		templateHTML2 = templateHTML2 &  "<option value=""None"">No Year Filter</option>"
	End If
	If FilterYear <> "None" And FilterYear <> "Use Year Filter" Then
		templateHTML2 = templateHTML2 &  "<option value=""None"">No Year Filter</option>"
		templateHTML2 = templateHTML2 &  "<option value=""Use Year Filter"">Use Year Filter</option>"
	End If
	For i = 1 To YearList.Count - 1
		If FilterYear <> YearList.Item(i) Or FilterYear = "None" Or FilterYear = "Use Year Filter" Then
			templateHTML2 = templateHTML2 &  "<option value=""" & EncodeHtmlChars(YearList.Item(i)) & """>" & YearList.Item(i) & "</option>"
		Else
			templateHTML2 = templateHTML2 &  "<option value=""" & EncodeHtmlChars(YearList.Item(i)) & """ selected>" & YearList.Item(i) & "</option>"
		End If
	Next
	
	templateHTML2 = templateHTML2 &  "</select>"
	templateHTML2 = templateHTML2 &  "</td>"
	'Filters End
	templateHTML2 = templateHTML2 &  "</tr>"
	templateHTML2 = templateHTML2 &  "</table>"
	templateHTML2 = templateHTML2 &  "</td>"
	templateHTML2 = templateHTML2 &  "</tr>"
	
	GetHeader = templateHTML2
	
End Function


Function GetFooter()
	
	Dim templateHTML2
	templateHTML2 = "</table>"
	templateHTML2 = templateHTML2 &  "</body>"
	templateHTML2 = templateHTML2 &  "</HTML>"
	
	GetFooter = templateHTML2
	
End Function



' We use this procedure to reformat results as soon as they are downloaded
Sub FormatSearchResultsViewer(Tracks, TracksNum, TracksCD, Durations, AlbumArtist, AlbumArtistTitle, ArtistTitles, AlbumTitle, ReleaseDate, OriginalDate, Genres, Styles, theLabels, theCountry, theArt, releaseID, Catalog, Lyricists, Composers, Conductors, Producers, InvolvedPeople, theFormat, theMaster, comment, DiscogsTracksNum, DataQuality)
	
	WriteLog "Start FormatSearchResultsViewer"
	Dim templateHTML, checkBox, text, listBox, submitButton
	Dim SelectedTracksCount, UnSelectedTracksCount
	Dim SubTrackFlag
	Dim i, theTracks, currentCD, theGenres
	templateHTML = GetHeader()
	
	' Titles Begin
	templateHTML = templateHTML &  "<tr>"
	templateHTML = templateHTML &  "<td align=left bgcolor=""#CCCCCC""><b>Album Art:</b></td>"
	templateHTML = templateHTML &  "<td align=left bgcolor=""#CCCCCC""><b>Release Information:</b></td>"
	templateHTML = templateHTML &  "<td align=left bgcolor=""#CCCCCC""><b>Tracklisting:</b></td>"
	templateHTML = templateHTML &  "</tr>"
	' Titles End
	templateHTML = templateHTML &  "<tr>"
	' Release Cover Begin
	templateHTML = templateHTML &  "<td align=left valign=top>"
	templateHTML = templateHTML &  "<table border=0 cellspacing=0 cellpadding=1 class=tabletext>"
	If theArt <> "" Then
		templateHTML = templateHTML &  "<tr><td colspan=2><a href=""http://www.discogs.com/viewimages?release=<!RELEASEID!>"" target=""_blank""><img src=""<!COVER!>"" border=""0""/></a></td></tr>"
	Else
		templateHTML = templateHTML &  "<tr><td colspan=2><table width=150 height=150 border=1><tr><td><center>No Image<br>Available</center></td></tr></table></td></tr>"
	End If
	templateHTML = templateHTML &  "<tr><td colspan=2 align=center><input type=checkbox id=""cover"" >Large <input type=checkbox id=""smallcover"" >Small (150px)</td></tr>"
	templateHTML = templateHTML &  "<tr><td colspan=2 align=center><br></td></tr>"
	' Release Cover End
	
	' Options Begin
	templateHTML = templateHTML &  "<tr><td colspan=2 align=center><button type=button class=tabletext id=""saveoptions"">Save Options</button></td></tr>"
	templateHTML = templateHTML &  "<tr><td align=center colspan=2><b>Options:</b></td></tr>"
	templateHTML = templateHTML &  "<tr><td colspan=2 align=left><input type=checkbox id=""lyricist"" >Save Lyricist</td></tr>"
	templateHTML = templateHTML &  "<tr><td colspan=2 align=left><input type=checkbox id=""composer"" >Save Composer</td></tr>"
	templateHTML = templateHTML &  "<tr><td colspan=2 align=left><input type=checkbox id=""conductor"" >Save Conductor</td></tr>"
	templateHTML = templateHTML &  "<tr><td colspan=2 align=left><input type=checkbox id=""producer"" >Save Producer</td></tr>"
	templateHTML = templateHTML &  "<tr><td colspan=2 align=left><input type=checkbox id=""involved"" >Save Involved People</td></tr>"
	templateHTML = templateHTML &  "<tr><td colspan=2 align=left><input type=checkbox id=""comments"" >Save Comment</td></tr>"
	templateHTML = templateHTML &  "<tr><td colspan=2 align=left><input type=checkbox id=""useanv"" title=""Artist Name Variation - Using no name variation (e.g. nickname)"" >Don't Use ANV's</td></tr>"
	templateHTML = templateHTML &  "<tr><td colspan=2 align=left><input type=checkbox id=""yearonlydate"" title=""If checked only the Year will be saved (e.g. 14.01.1982 -> 1982)"" >Only Year Of Date</td></tr>"
	templateHTML = templateHTML &  "<tr><td colspan=2 align=left><input type=checkbox id=""titlefeaturing"" title=""If checked the feat. Artist appears in the title tag (e.g. Aaliyah (ft. Timbaland) - We Need a Resolution  ->  Aaliyah - We Need a Resolution (ft. Timbaland) )"" >feat. Artist behind Title</td></tr>"
	
	templateHTML = templateHTML &  "<tr><td colspan=2 align=left><input type=checkbox id=""FeaturingName"" title=""Rename 'feat.' to the given word"" >Rename 'feat.' to:</td></tr>"
	templateHTML = templateHTML &  "<tr><td colspan=2 align=left><input type=text id=""TxtFeaturingName"" ></td></tr>"
	
	templateHTML = templateHTML &  "<tr><td colspan=2 align=left><input type=checkbox id=""various"" title=""Rename 'Various' Artist to the given word"" >Rename 'Various' Artist to:</td></tr>"
	templateHTML = templateHTML &  "<tr><td colspan=2 align=left><input type=text id=""txtvarious"" ></td></tr>"
	
	templateHTML = templateHTML &  "<tr><td align=center colspan=2><br></td></tr>"
	templateHTML = templateHTML &  "<tr><td align=center colspan=2><b>Disc/Track Numbering:</b></td></tr>"
	templateHTML = templateHTML &  "<tr><td colspan=2 align=left><input type=checkbox id=""UnselectNoTrackPos"" title=""Tracks without track-number at discogs will automatically unselect (Info-Tracks e.g. 'Bonus Tracks')"" >Unselect Track without Track-number</td></tr>"
	templateHTML = templateHTML &  "<tr><td colspan=2 align=left><input type=checkbox id=""SubTrackNameSelection"" title=""If checked the Sub-Track will be named like 'Sub-Track 1, Sub-Track 2, Sub Track 3'  if not checked the Sub-Tracks will be named like 'Track Name (Sub-Track 1, Sub-Track 2, Sub Track 3)'"" >Other Sub-Track Naming</td></tr>"
	
	templateHTML = templateHTML &  "<tr><td colspan=2 align=left><input type=checkbox id=""forcenumeric"" >Force To Numeric</td></tr>"
	templateHTML = templateHTML &  "<tr><td colspan=2 align=left><input type=checkbox id=""sidestodisc"" >Sides To Disc</td></tr>"
	templateHTML = templateHTML &  "<tr><td colspan=2 align=left><input type=checkbox id=""forcedisc"" >Force Disc Usage</td></tr>"
	templateHTML = templateHTML &  "<tr><td colspan=2 align=left><input type=checkbox id=""nodisc"" title=""Prevent the script from interpret sub tracks as disc-numbers"" >Force NO Disc Usage</td></tr>"
	templateHTML = templateHTML &  "<tr><td colspan=2 align=left><input type=checkbox id=""leadingzero"" >Add Leading Zero</td></tr>"
	templateHTML = templateHTML &  "<tr><td colspan=2 align=left><input type=checkbox id=""SkipNotChangedReleases"" title=""Skip unchanged releases automatically and check the next one"" >Skip unchanged releases</td></tr>"
	templateHTML = templateHTML &  "<tr><td colspan=2 align=left><input type=checkbox id=""ProcessOnlyDiscogs"" title=""Process only albums already found at discogs"" >Process only Discogs Releases</td></tr>"
	templateHTML = templateHTML &  "<tr><td colspan=2 align=left><input type=checkbox id=""ProcessNoDiscogs"" title=""Process only albums not found at discogs"" >Process no Discogs releases</td></tr>"
	templateHTML = templateHTML &  "<tr><td colspan=2 align=center><br></td></tr>"
	
	templateHTML = templateHTML &  "</table>"
	templateHTML = templateHTML &  "</td>"
	' Options End
	
	' Release Information Begin
	templateHTML = templateHTML &  "<td align=left valign=top>"
	templateHTML = templateHTML &  "<table border=0 cellspacing=0 cellpadding=1 class=tabletext>"

	theTracks = ""
	currentCD = 0
	iMaxTracks = Tracks.Count
	If TracksCD.Count < iMaxTracks Then
		iMaxTracks = TracksCD.Count
	End If
	
	'Check for different Track number
	SelectedTracksCount = 0
	UnSelectedTracksCount = 0
	SubTrackFlag = False
	For i = 0 To iMaxTracks - 1
		If (UnselectedTracks(i) = "") Then
			If InStr(DiscogsTracksNum.Item(i), ".") <> 0 Then
				If SubTrackFlag = False Then
					SubTrackFlag = True
					SelectedTracksCount = SelectedTracksCount + 1
				End If
			Else
				If SubTrackFlag = True Then SubTrackFlag = False
				SelectedTracksCount = SelectedTracksCount + 1
			End If
		Else
			UnSelectedTracksCount = UnSelectedTracksCount + 1
		End If
	Next
	
	If (iMaxTracks - UnSelectedTracksCount) <> NewTrackList.Count Then
		templateHTML = templateHTML &  "<tr><td colspan=3 align=center><b><span style=""color:#FF0000"">There are different numbers of tracks !</span></b></td></tr>"
		templateHTML = templateHTML &  "<tr><td colspan=3 align=center><br></td></tr>"
	End If
	
	templateHTML = templateHTML &  "<tr>"
	templateHTML = templateHTML &  "<td><input type=checkbox id=""releaseid"" ></td>"
	templateHTML = templateHTML &  "<td>Release:</td>"
	If (theMaster <> "") Then
		templateHTML = templateHTML &  "<td><a href=""http://www.discogs.com/release/<!RELEASEID!>"" target=""_blank""><!RELEASEID!></a> (Master: <a href=""http://www.discogs.com/master/<!MASTERID!>"" target=""_blank""><!MASTERID!></a>)</td>"
	Else
		templateHTML = templateHTML &  "<td><a href=""http://www.discogs.com/release/<!RELEASEID!>"" target=""_blank""><!RELEASEID!></a> (Master: N/A)</td>"
	End If
	templateHTML = templateHTML &  "</tr>"
	templateHTML = templateHTML &  "<tr>"
	templateHTML = templateHTML &  "<td><input type=checkbox id=""artist"" ></td>"
	templateHTML = templateHTML &  "<td>Artist:</td>"
	templateHTML = templateHTML &  "<td><a href=""http://www.discogs.com/artist/<!ARTIST!>"" target=""_blank""><!ARTIST!></a></td>"
	templateHTML = templateHTML &  "</tr>"
	templateHTML = templateHTML &  "<tr>"
	templateHTML = templateHTML &  "<td><input type=checkbox id=""album"" ></td>"
	templateHTML = templateHTML &  "<td>Album:</td>"
	templateHTML = templateHTML &  "<td><a href=""http://www.discogs.com/release/<!RELEASEID!>"" target=""_blank""><!ALBUMTITLE!></a></td>"
	templateHTML = templateHTML &  "</tr>"
	templateHTML = templateHTML &  "<tr>"
	templateHTML = templateHTML &  "<td><input type=checkbox id=""albumartist"" ><input type=checkbox id=""albumartistfirst"" ></td>"
	templateHTML = templateHTML &  "<td>Album Artist:</td>"
	templateHTML = templateHTML &  "<td><a href=""http://www.discogs.com/artist/<!ALBUMARTIST!>"" target=""_blank""><!ALBUMARTIST!></a></td>"
	templateHTML = templateHTML &  "</tr>"
	templateHTML = templateHTML &  "<tr>"
	templateHTML = templateHTML &  "<td><input type=checkbox id=""label"" ></td>"
	templateHTML = templateHTML &  "<td>Label:</td>"
	templateHTML = templateHTML &  "<td><a href=""http://www.discogs.com/label/<!LABEL!>"" target=""_blank""><!LABEL!></a></td>"
	templateHTML = templateHTML &  "</tr>"
	templateHTML = templateHTML &  "<tr>"
	templateHTML = templateHTML &  "<td><input type=checkbox id=""catalog"" ></td>"
	templateHTML = templateHTML &  "<td>Catalog#:</td>"
	templateHTML = templateHTML &  "<td><!CATALOG!></td>"
	templateHTML = templateHTML &  "</tr>"
	templateHTML = templateHTML &  "<tr>"
	templateHTML = templateHTML &  "<td><input type=checkbox id=""format"" ></td>"
	templateHTML = templateHTML &  "<td>Format:</td>"
	templateHTML = templateHTML &  "<td><!FORMAT!></td>"
	templateHTML = templateHTML &  "</tr>"
	templateHTML = templateHTML &  "<tr>"
	templateHTML = templateHTML &  "<td><input type=checkbox id=""country"" ></td>"
	templateHTML = templateHTML &  "<td>Country:</td>"
	templateHTML = templateHTML &  "<td><!COUNTRY!></td>"
	templateHTML = templateHTML &  "</tr>"
	templateHTML = templateHTML &  "<tr>"
	templateHTML = templateHTML &  "<td><input type=checkbox title=""If option set, the release date of this Discogs release will be saved"" id=""date"" ></td>"
	templateHTML = templateHTML &  "<td>Date:</td>"
	templateHTML = templateHTML &  "<td><!RELEASEDATE!></td>"
	templateHTML = templateHTML &  "</tr>"
	templateHTML = templateHTML &  "<tr>"
	templateHTML = templateHTML &  "<td><input type=checkbox title=""If option set, the release date of the Discogs master release will be saved"" id=""origdate"" ></td>"
	templateHTML = templateHTML &  "<td>Original Date:</td>"
	templateHTML = templateHTML &  "<td><!ORIGDATE!></td>"
	templateHTML = templateHTML &  "</tr>"
	templateHTML = templateHTML &  "<tr>"
	templateHTML = templateHTML &  "<td><input type=checkbox id=""genre"" ><input type=checkbox id=""style"" ></td>"
	templateHTML = templateHTML &  "<td>Genre:</td>"
	templateHTML = templateHTML &  "<td><!GENRE!></td>"
	templateHTML = templateHTML &  "</tr>"
	templateHTML = templateHTML &  "<tr>"
	templateHTML = templateHTML &  "<td colspan=2>Release Data Quality:</td>"
	templateHTML = templateHTML &  "<td><!DATAQUALITY!></td>"
	templateHTML = templateHTML &  "</tr>"
	templateHTML = templateHTML &  "</table>"
	templateHTML = templateHTML &  "</td>"
	' Release Information End
	' Tracklisting Begin
	templateHTML = templateHTML & "<td align=left valign=top>"
	templateHTML = templateHTML & "<table border=0 cellspacing=0 cellpadding=1 class=tabletext>"
	templateHTML = templateHTML & "<tr>"
	If CheckOriginalDiscogsTrack Then
		templateHTML = templateHTML & "<td align=left><b>Discogs</b></td>"
	Else
		templateHTML = templateHTML & "<td> </td>"
	End If
	templateHTML = templateHTML & "<td><input type=checkbox id=""selectall""></td>"
	templateHTML = templateHTML & "<td align=center><input type=checkbox id=""discnum""></td>"
	templateHTML = templateHTML & "<td align=center><input type=checkbox id=""tracknum""></td>"
	templateHTML = templateHTML & "<td align=right><b>Artist</b></td>"
	templateHTML = templateHTML & "<td> </td>"
	templateHTML = templateHTML & "<td align=left><b>Title</b></td>"
	templateHTML = templateHTML & "<td align=right><b>Duration</b></td>"
	templateHTML = templateHTML & "</tr>"
	
	For i=0 To iMaxTracks - 1
		templateHTML = templateHTML &  "<tr>"
		If CheckOriginalDiscogsTrack Then
			templateHTML = templateHTML & "<td align=center>" & DiscogsTracksNum.Item(i) & "</td>"
		Else
			templateHTML = templateHTML & "<td> </td>"
		End If
		If(UnselectedTracks(i) = "") Then
			templateHTML = templateHTML & "<td><input type=checkbox id=""unselected["&i&"]"" checked></td>"
		Else
			templateHTML = templateHTML & "<td><input type=checkbox id=""unselected["&i&"]""></td>"
		End If
		templateHTML = templateHTML & "<td align=center>" & TracksCD.Item(i) & "</td>"
		templateHTML = templateHTML & "<td align=center>" & TracksNum.Item(i) & "</td>"
		templateHTML = templateHTML & "<td align=right><b>" & ArtistTitles.Item(i) & "</b></td>"
		templateHTML = templateHTML & "<td align=center><b>-</b></td>"
		templateHTML = templateHTML & "<td align=left><b>" & Tracks.Item(i) & "</b></td>"
		templateHTML = templateHTML & "<td align=right>" & Durations.Item(i) & "</td>"
		templateHTML = templateHTML & "</tr>"
		If(CheckLyricist And Lyricists.Item(i) <> "") Then templateHTML = templateHTML &  "<tr><td colspan=5></td><td colspan=2 align=left>Lyrics: "& Lyricists.Item(i) &"</td></tr>"
		If(CheckComposer And Composers.Item(i) <> "") Then templateHTML = templateHTML &  "<tr><td colspan=5></td><td colspan=2 align=left>Composer: "& Composers.Item(i) &"</td></tr>"
		If(CheckConductor And Conductors.Item(i) <> "") Then templateHTML = templateHTML &  "<tr><td colspan=5></td><td colspan=2 align=left>Conductor: "& Conductors.Item(i) &"</td></tr>"
		If(CheckProducer And Producers.Item(i) <> "") Then templateHTML = templateHTML &  "<tr><td colspan=5></td><td colspan=2 align=left>Producer: "& Producers.Item(i) &"</td></tr>"
		
		If(CheckInvolved And InvolvedPeople.Item(i) <> "") Then templateHTML = templateHTML &  "<tr><td colspan=5></td><td colspan=2 align=left>"& InvolvedPeople.Item(i) &"</td></tr>"
	Next
	
	templateHTML = templateHTML &  "</table>"
	templateHTML = templateHTML &  "</td>"
	' Tracklisting End
	
	templateHTML = templateHTML &  GetFooter()
	
	
	templateHTML = Replace(templateHTML, "<!RELEASEID!>", releaseID)
	templateHTML = Replace(templateHTML, "<!MASTERID!>", theMaster)
	templateHTML = Replace(templateHTML, "<!ARTIST!>", AlbumArtistTitle)
	templateHTML = Replace(templateHTML, "<!ALBUMARTIST!>",  AlbumArtist)
	templateHTML = Replace(templateHTML, "<!ALBUMTITLE!>", AlbumTitle)
	templateHTML = Replace(templateHTML, "<!RELEASEDATE!>", ReleaseDate)
	templateHTML = Replace(templateHTML, "<!ORIGDATE!>", OriginalDate)
	templateHTML = Replace(templateHTML, "<!LABEL!>", theLabels)
	templateHTML = Replace(templateHTML, "<!COUNTRY!>", theCountry)
	templateHTML = Replace(templateHTML, "<!COVER!>", theArt)
	templateHTML = Replace(templateHTML, "<!CATALOG!>", Catalog)
	templateHTML = Replace(templateHTML, "<!FORMAT!>", theFormat)
	templateHTML = Replace(templateHTML, "<!DATAQUALITY!>", DataQuality)

	theGenres = ""
	
	If Genres <> "" Then
		If CheckGenre Then
			theGenres = Genres
		Else
			theGenres = "<s>" + Genres + "</s>"
		End If
	End If
	
	If Styles <> "" Then
		If theGenres <> "" Then
			If CheckGenre Then
				theGenres = theGenres & Separator
			Else
				theGenres = theGenres & "<s>" & Separator & "</s>"
			End If
		End If
		If CheckStyle Then
			theGenres = theGenres & Styles
		Else
			theGenres = theGenres & "<s>" & Styles & "</s>"
		End If
	End If
	templateHTML = Replace(templateHTML, "<!GENRE!>", theGenres)
	
	WebBrowser.SetHTMLDocument templateHTML

	Dim templateHTMLDoc
	Set templateHTMLDoc = WebBrowser.Interf.Document

	Set checkBox = templateHTMLDoc.getElementById("album")
	checkBox.Checked = CheckAlbum
	Script.RegisterEvent checkBox, "onclick", "Update"
	Set checkBox = templateHTMLDoc.getElementById("artist")
	checkBox.Checked = CheckArtist
	Script.RegisterEvent checkBox, "onclick", "Update"
	Set checkBox = templateHTMLDoc.getElementById("albumartist")
	checkBox.Checked = CheckAlbumArtist
	Script.RegisterEvent checkBox, "onclick", "Update"
	Set checkBox = templateHTMLDoc.getElementById("albumartistfirst")
	checkBox.Checked = CheckAlbumArtistFirst
	Script.RegisterEvent checkBox, "onclick", "Update"
	Set checkBox = templateHTMLDoc.getElementById("date")
	checkBox.Checked = CheckDate
	Script.RegisterEvent checkBox, "onclick", "Update"
	Set checkBox = templateHTMLDoc.getElementById("origdate")
	checkBox.Checked = CheckOrigDate
	Script.RegisterEvent checkBox, "onclick", "Update"
	Set checkBox = templateHTMLDoc.getElementById("label")
	checkBox.Checked = CheckLabel
	Script.RegisterEvent checkBox, "onclick", "Update"
	Set checkBox = templateHTMLDoc.getElementById("country")
	checkBox.Checked = CheckCountry
	Script.RegisterEvent checkBox, "onclick", "Update"
	Set checkBox = templateHTMLDoc.getElementById("genre")
	checkBox.Checked = CheckGenre
	Script.RegisterEvent checkBox, "onclick", "Update"
	Set checkBox = templateHTMLDoc.getElementById("style")
	checkBox.Checked = CheckStyle
	Script.RegisterEvent checkBox, "onclick", "Update"
	Set checkBox = templateHTMLDoc.getElementById("cover")
	checkBox.Checked = CheckCover
	Script.RegisterEvent checkBox, "onclick", "Update"
	Set checkBox = templateHTMLDoc.getElementById("smallcover")
	checkBox.Checked = CheckSmallCover
	Script.RegisterEvent checkBox, "onclick", "Update"
	Set checkBox = templateHTMLDoc.getElementById("catalog")
	checkBox.Checked = CheckCatalog
	Script.RegisterEvent checkBox, "onclick", "Update"
	Set checkBox = templateHTMLDoc.getElementById("releaseid")
	checkBox.Checked = CheckRelease
	Script.RegisterEvent checkBox, "onclick", "Update"
	Set checkBox = templateHTMLDoc.getElementById("involved")
	checkBox.Checked = CheckInvolved
	Script.RegisterEvent checkBox, "onclick", "Update"
	Set checkBox = templateHTMLDoc.getElementById("lyricist")
	checkBox.Checked = CheckLyricist
	Script.RegisterEvent checkBox, "onclick", "Update"
	Set checkBox = templateHTMLDoc.getElementById("composer")
	checkBox.Checked = CheckComposer
	Script.RegisterEvent checkBox, "onclick", "Update"
	Set checkBox = templateHTMLDoc.getElementById("conductor")
	checkBox.Checked = CheckConductor
	Script.RegisterEvent checkBox, "onclick", "Update"
	Set checkBox = templateHTMLDoc.getElementById("producer")
	checkBox.Checked = CheckProducer
	Script.RegisterEvent checkBox, "onclick", "Update"
	Set checkBox = templateHTMLDoc.getElementById("discnum")
	checkBox.Checked = CheckDiscNum
	Script.RegisterEvent checkBox, "onclick", "Update"
	Set checkBox = templateHTMLDoc.getElementById("tracknum")
	checkBox.Checked = CheckTrackNum
	Script.RegisterEvent checkBox, "onclick", "Update"
	Set checkBox = templateHTMLDoc.getElementById("format")
	checkBox.Checked = CheckFormat
	Script.RegisterEvent checkBox, "onclick", "Update"
	Set checkBox = templateHTMLDoc.getElementById("useanv")
	checkBox.Checked = CheckUseAnv
	Script.RegisterEvent checkBox, "onclick", "Update"
	Set checkBox = templateHTMLDoc.getElementById("yearonlydate")
	checkBox.Checked = CheckYearOnlyDate
	Script.RegisterEvent checkBox, "onclick", "Update"
	Set checkBox = templateHTMLDoc.getElementById("forcenumeric")
	checkBox.Checked = CheckForceNumeric
	Script.RegisterEvent checkBox, "onclick", "Update"
	Set checkBox = templateHTMLDoc.getElementById("sidestodisc")
	checkBox.Checked = CheckSidesToDisc
	Script.RegisterEvent checkBox, "onclick", "Update"
	Set checkBox = templateHTMLDoc.getElementById("forcedisc")
	checkBox.Checked = CheckForceDisc
	Script.RegisterEvent checkBox, "onclick", "Update"
	Set checkBox = templateHTMLDoc.getElementById("nodisc")
	checkBox.Checked = CheckNoDisc
	Script.RegisterEvent checkBox, "onclick", "Update"
	Set checkBox = templateHTMLDoc.getElementById("leadingzero")
	checkBox.Checked = CheckLeadingZero
	Script.RegisterEvent checkBox, "onclick", "Update"
	Set checkBox = templateHTMLDoc.getElementById("titlefeaturing")
	checkBox.Checked = CheckTitleFeaturing
	Script.RegisterEvent checkBox, "onclick", "Update"
	Set text = templateHTMLDoc.getElementById("TxtFeaturingName")
	text.value = TxtFeaturingName
	Script.RegisterEvent text, "onchange", "Update"
	Set checkbox = templateHTMLDoc.getElementById("FeaturingName")
	checkBox.Checked = CheckFeaturingName
	Script.RegisterEvent checkBox, "onclick", "Update"
	Set checkBox = templateHTMLDoc.getElementById("comments")
	checkBox.Checked = CheckComment
	Script.RegisterEvent checkBox, "onclick", "Update"
	Set text = templateHTMLDoc.getElementById("txtvarious")
	text.value = TxtVarious
	Script.RegisterEvent text, "onchange", "Update"
	Set checkBox = templateHTMLDoc.getElementById("various")
	checkBox.Checked = CheckVarious
	Script.RegisterEvent checkBox, "onclick", "Update"
	Set checkBox = templateHTMLDoc.getElementById("UnselectNoTrackPos")
	checkBox.Checked = CheckUnselectNoTrackPos
	Script.RegisterEvent checkBox, "onclick", "NoTrackPos"
	Set checkBox = templateHTMLDoc.getElementById("SubTrackNameSelection")
	checkBox.Checked = SubTrackNameSelection
	Script.RegisterEvent checkBox, "onclick", "Update"

	Set checkBox = templateHTMLDoc.getElementById("SkipNotChangedReleases")
	checkBox.Checked = SkipNotChangedReleases
	Script.RegisterEvent checkBox, "onclick", "Update"

	Set listBox = templateHTMLDoc.getElementById("filtermediatype")
	Script.RegisterEvent listBox, "onchange", "Filter"
	Set listBox = templateHTMLDoc.getElementById("filtermediaformat")
	Script.RegisterEvent listBox, "onchange", "Filter"
	Set listBox = templateHTMLDoc.getElementById("filtercountry")
	Script.RegisterEvent listBox, "onchange", "Filter"
	Set listBox = templateHTMLDoc.getElementById("filteryear")
	Script.RegisterEvent listBox, "onchange", "Filter"
	Set listBox = templateHTMLDoc.getElementById("load")
	Script.RegisterEvent listBox, "onchange", "Filter"

	Set checkBox = templateHTMLDoc.getElementById("selectall")
	checkBox.Checked = SelectAll
	Script.RegisterEvent checkBox, "onclick", "SwitchAll"

	Set listBox = templateHTMLDoc.getElementById("alternative")
	Script.RegisterEvent listBox, "onchange", "Alternative"

	Set submitButton = templateHTMLDoc.getElementById("saveoptions")
	Script.RegisterEvent submitButton, "onclick", "SaveOptions"

	Set submitButton = templateHTMLDoc.getElementById("showcountryfilter")
	Script.RegisterEvent submitButton, "onclick", "ShowCountryFilter"

	Set submitButton = templateHTMLDoc.getElementById("showmediatypefilter")
	Script.RegisterEvent submitButton, "onclick", "ShowMediaTypeFilter"
	
	Set submitButton = templateHTMLDoc.getElementById("showmediaformatfilter")
	Script.RegisterEvent submitButton, "onclick", "ShowMediaFormatFilter"
	
	Set submitButton = templateHTMLDoc.getElementById("showyearfilter")
	Script.RegisterEvent submitButton, "onclick", "ShowYearFilter"
	
	Set checkBox = templateHTMLDoc.getElementById("ProcessOnlyDiscogs")
	checkBox.checked = ProcessOnlyDiscogs
	Script.RegisterEvent checkBox, "onclick", "Update_ProcessOnlyDiscogs"
	
	Set checkBox = templateHTMLDoc.getElementById("ProcessNoDiscogs")
	checkBox.checked = ProcessNoDiscogs
	Script.RegisterEvent checkBox, "onclick", "Update_ProcessNoDiscogs"
	
	WriteLog "Stop FormatSearchResultsViewer"
	
End Sub


Sub Update()
	
	Dim templateHTMLDoc, checkBox, text
	Set WebBrowser = SDB.Objects("WebBrowser")
	Set templateHTMLDoc = WebBrowser.Interf.Document

	Set checkBox = templateHTMLDoc.getElementById("album")
	CheckAlbum = checkBox.Checked
	Set checkBox = templateHTMLDoc.getElementById("artist")
	CheckArtist = checkBox.Checked
	Set checkBox = templateHTMLDoc.getElementById("albumartist")
	CheckAlbumArtist = checkBox.Checked
	Set checkBox = templateHTMLDoc.getElementById("albumartistfirst")
	CheckAlbumArtistFirst = checkBox.Checked
	Set checkBox = templateHTMLDoc.getElementById("date")
	CheckDate = checkBox.Checked
	Set checkBox = templateHTMLDoc.getElementById("origdate")
	CheckOrigDate = checkBox.Checked
	Set checkBox = templateHTMLDoc.getElementById("label")
	CheckLabel = checkBox.Checked
	Set checkBox = templateHTMLDoc.getElementById("genre")
	CheckGenre = checkBox.Checked
	Set checkBox = templateHTMLDoc.getElementById("style")
	CheckStyle = checkBox.Checked
	Set checkBox = templateHTMLDoc.getElementById("country")
	CheckCountry = checkBox.Checked
	Set checkBox = templateHTMLDoc.getElementById("cover")
	If Not CheckCover And checkBox.Checked Then
		CheckSmallCover = False
		CheckCover = checkBox.Checked
	Else
		CheckCover = checkBox.Checked
		Set checkBox = templateHTMLDoc.getElementById("smallcover")
		If Not CheckSmallCover And checkBox.Checked Then
			CheckCover = False
		End If
		CheckSmallCover = checkBox.Checked
	End If
	Set checkBox = templateHTMLDoc.getElementById("catalog")
	CheckCatalog = checkBox.Checked
	Set checkBox = templateHTMLDoc.getElementById("releaseid")
	CheckRelease = checkBox.Checked
	Set checkBox = templateHTMLDoc.getElementById("involved")
	CheckInvolved = checkBox.Checked
	Set checkBox = templateHTMLDoc.getElementById("lyricist")
	CheckLyricist = checkBox.Checked
	Set checkBox = templateHTMLDoc.getElementById("composer")
	CheckComposer = checkBox.Checked
	Set checkBox = templateHTMLDoc.getElementById("conductor")
	CheckConductor = checkBox.Checked
	Set checkBox = templateHTMLDoc.getElementById("producer")
	CheckProducer = checkBox.Checked
	Set checkBox = templateHTMLDoc.getElementById("discnum")
	CheckDiscNum = checkBox.Checked
	Set checkBox = templateHTMLDoc.getElementById("tracknum")
	CheckTrackNum = checkBox.Checked
	Set checkBox = templateHTMLDoc.getElementById("format")
	CheckFormat = checkBox.Checked
	Set checkBox = templateHTMLDoc.getElementById("useanv")
	CheckUseAnv = checkBox.Checked
	Set checkBox = templateHTMLDoc.getElementById("yearonlydate")
	CheckYearOnlyDate = checkBox.Checked
	Set checkBox = templateHTMLDoc.getElementById("forcenumeric")
	CheckForceNumeric = checkBox.Checked
	Set checkBox = templateHTMLDoc.getElementById("sidestodisc")
	CheckSidesToDisc = checkBox.Checked
	Set checkBox = templateHTMLDoc.getElementById("forcedisc")
	If Not CheckForceDisc And checkBox.Checked Then
		CheckNoDisc = False
		CheckForceDisc = checkBox.Checked
	Else
		CheckForceDisc = checkBox.Checked
		Set checkBox = templateHTMLDoc.getElementById("nodisc")
		If Not CheckNoDisc And checkBox.Checked Then
			CheckForceDisc = False
		End If
		CheckNoDisc = checkBox.Checked
	End If
	Set checkBox = templateHTMLDoc.getElementById("leadingzero")
	CheckLeadingZero = checkBox.Checked
	Set checkBox = templateHTMLDoc.getElementById("titlefeaturing")
	CheckTitleFeaturing = checkBox.Checked
	Set checkBox = templateHTMLDoc.getElementById("FeaturingName")
	CheckFeaturingName = checkBox.Checked
	Set text = templateHTMLDoc.getElementById("TxtFeaturingName")
	TxtFeaturingName = text.Value
	Set checkBox = templateHTMLDoc.getElementById("comments")
	CheckComment = checkBox.Checked
	Set checkBox = templateHTMLDoc.getElementById("various")
	CheckVarious = checkBox.Checked
	Set text = templateHTMLDoc.getElementById("txtvarious")
	TxtVarious = text.Value
	Set checkBox = templateHTMLDoc.getElementById("SubTrackNameSelection")
	SubTrackNameSelection = checkBox.Checked
	
	Set checkBox = templateHTMLDoc.getElementById("SkipNotChangedReleases")
	SkipNotChangedReleases = checkBox.Checked
	
	ReloadResults
	
End Sub


Sub Update_ProcessOnlyDiscogs()

	Dim checkBox, templateHTMLDoc
	Set WebBrowser = SDB.Objects("WebBrowser")
	Set templateHTMLDoc = WebBrowser.Interf.Document
	Set checkBox = templateHTMLDoc.getElementById("ProcessOnlyDiscogs")
	ProcessOnlyDiscogs = checkBox.Checked
	If ProcessOnlyDiscogs = True Then
		Set checkBox = templateHTMLDoc.getElementById("ProcessNoDiscogs")
		ProcessNoDiscogs = False
		checkBox.Checked= False
	End If
	
End Sub



Sub Update_ProcessNoDiscogs()

	Dim checkBox, templateHTMLDoc
	Set WebBrowser = SDB.Objects("WebBrowser")
	Set checkBox = templateHTMLDoc.getElementById("ProcessNoDiscogs")
	ProcessNoDiscogs = checkBox.Checked
	If ProcessNoDiscogs = True Then
		Set checkBox = templateHTMLDoc.getElementById("ProcessOnlyDiscogs")
		ProcessOnlyDiscogs = False
		checkBox.Checked= False
	End If
	
End Sub

Sub NoTrackPos()

	Dim checkBox, templateHTMLDoc, i
	Set WebBrowser = SDB.Objects("WebBrowser")
	Set templateHTMLDoc = WebBrowser.Interf.Document
	Set checkBox = templateHTMLDoc.getElementById("UnselectNoTrackPos")
	CheckUnselectNoTrackPos = checkBox.Checked
	If Not CheckUnselectNoTrackPos Then
		For i = 0 To iMaxTracks - 1
			UnselectedTracks(i) = ""
		Next
	End If
	
	ReloadResults
	
End Sub


Sub SwitchAll()

	Dim templateHTMLDoc, i, checkBox
	Set WebBrowser = SDB.Objects("WebBrowser")
	Set templateHTMLDoc = WebBrowser.Interf.Document
	Set checkBox = templateHTMLDoc.getElementById("selectall")
	SelectAll = checkBox.Checked
	UserChoose = True
	
	For i = 0 To iMaxTracks - 1
		If SelectAll Then
			UnselectedTracks(i) = ""
		Else
			UnselectedTracks(i) = "x"
		End If
	Next
	
	ReloadResults
	
End Sub


Sub Filter()

	WriteLog "Start Filter"
	Dim templateHTMLDoc, listBox
	Set WebBrowser = SDB.Objects("WebBrowser")
	Set templateHTMLDoc = WebBrowser.Interf.Document
	
	Set listBox = templateHTMLDoc.getElementById("filtermediatype")
	FilterMediaType = listBox.Value
	If FilterMediaType = "None" Then
		MediaTypeFilterList.Item(0) = "0"
	ElseIf FilterMediaType = "Use MediaType Filter" Then
		MediaTypeFilterList.Item(0) = "1"
	Else
		MediaTypeFilterList.Item(0) = FilterMediaType
	End If
	Set listBox = templateHTMLDoc.getElementById("filtermediaformat")
	FilterMediaFormat = listBox.Value
	If FilterMediaFormat = "None" Then
		MediaFormatFilterList.Item(0) = "0"
	ElseIf FilterMediaFormat = "Use MediaFormat Filter" Then
		MediaFormatFilterList.Item(0) = "1"
	Else
		MediaFormatFilterList.Item(0) = FilterMediaFormat
	End If
	Set listBox = templateHTMLDoc.getElementById("filtercountry")
	FilterCountry = listBox.Value
	If FilterCountry = "None" Then
		CountryFilterList.Item(0) = "0"
	ElseIf FilterCountry = "Use Country Filter" Then
		CountryFilterList.Item(0) = "1"
	Else
		CountryFilterList.Item(0) = FilterCountry
	End If
	Set listBox = templateHTMLDoc.getElementById("filteryear")
	FilterYear = listBox.Value
	If FilterYear = "None" Then
		YearFilterList.Item(0) = "0"
	ElseIf FilterYear = "Use Year Filter" Then
		YearFilterList.Item(0) = "1"
	Else
		YearFilterList.Item(0) = FilterYear
	End If
	
	Set listBox = templateHTMLDoc.getElementById("load")
	CurrentLoadType = listBox.Value
	
	If(CurrentLoadType = "Master Release") Then
		WriteLog "SavedMasterId=" & SavedMasterId
		LoadMasterResults(SavedMasterId)
	ElseIf(CurrentLoadType = "Releases of Artist") Then
		WriteLog "SavedArtistId=" & SavedArtistId
		LoadArtistResults(SavedArtistId)
	ElseIf(CurrentLoadType = "Releases of Label") Then
		WriteLog "SavedLabelId=" & SavedLabelId
		LoadLabelResults(SavedLabelId)
	Else
		WriteLog "SavedSearchTerm=" & SavedSearchTerm
		FindResults SavedSearchTerm, "", ""
	End If
	WriteLog "Stop Filter"
	
End Sub


Sub Alternative()
	
	WriteLog("Start Alternative")
	Dim templateHTMLDoc
	Set WebBrowser = SDB.Objects("WebBrowser")
	Set templateHTMLDoc = WebBrowser.Interf.Document
	SavedSearchTerm =  templateHTMLDoc.getElementById("alternative").Value
	CurrentLoadType = "Search Results"
	FindResults SavedSearchTerm, "", ""
	REM ReloadResults
	WriteLog("Stop Alternative")
	
End Sub


Sub SaveOptions()

	WriteLog "Start SaveOptions"
	Dim a,tmp
	' save options if ini exists
	If Not (ini Is Nothing) Then
		WriteLog "Writing SaveOptions"
		ini.BoolValue("DiscogsAutoTagWeb","CheckAlbum") = CheckAlbum
		ini.BoolValue("DiscogsAutoTagWeb","CheckArtist") = CheckArtist
		ini.BoolValue("DiscogsAutoTagWeb","CheckAlbumArtist") = CheckAlbumArtist
		ini.BoolValue("DiscogsAutoTagWeb","CheckAlbumArtistFirst") = CheckAlbumArtistFirst
		ini.BoolValue("DiscogsAutoTagWeb","CheckLabel") = CheckLabel
		ini.BoolValue("DiscogsAutoTagWeb","CheckDate") = CheckDate
		ini.BoolValue("DiscogsAutoTagWeb","CheckOrigDate") = CheckOrigDate
		ini.BoolValue("DiscogsAutoTagWeb","CheckGenre") = CheckGenre
		ini.BoolValue("DiscogsAutoTagWeb","CheckStyle") = CheckStyle
		ini.BoolValue("DiscogsAutoTagWeb","CheckCountry") = CheckCountry
		ini.BoolValue("DiscogsAutoTagWeb","CheckCover") = CheckCover
		ini.BoolValue("DiscogsAutoTagWeb","CheckSmallCover") = CheckSmallCover
		ini.BoolValue("DiscogsAutoTagWeb","CheckCatalog") = CheckCatalog
		ini.BoolValue("DiscogsAutoTagWeb","CheckRelease") = CheckRelease
		ini.BoolValue("DiscogsAutoTagWeb","CheckInvolved") = CheckInvolved
		ini.BoolValue("DiscogsAutoTagWeb","CheckLyricist") = CheckLyricist
		ini.BoolValue("DiscogsAutoTagWeb","CheckComposer") = CheckComposer
		ini.BoolValue("DiscogsAutoTagWeb","CheckConductor") = CheckConductor
		ini.BoolValue("DiscogsAutoTagWeb","CheckProducer") = CheckProducer
		ini.BoolValue("DiscogsAutoTagWeb","CheckDiscNum") = CheckDiscNum
		ini.BoolValue("DiscogsAutoTagWeb","CheckTrackNum") = CheckTrackNum
		ini.BoolValue("DiscogsAutoTagWeb","CheckFormat") = CheckFormat
		ini.BoolValue("DiscogsAutoTagWeb","CheckUseAnv") = CheckUseAnv
		ini.BoolValue("DiscogsAutoTagWeb","CheckYearOnlyDate") = CheckYearOnlyDate
		ini.BoolValue("DiscogsAutoTagWeb","CheckForceNumeric") = CheckForceNumeric
		ini.BoolValue("DiscogsAutoTagWeb","CheckSidesToDisc") = CheckSidesToDisc
		ini.BoolValue("DiscogsAutoTagWeb","CheckForceDisc") = CheckForceDisc
		ini.BoolValue("DiscogsAutoTagWeb","CheckNoDisc") = CheckNoDisc
		ini.BoolValue("DiscogsAutoTagWeb","CheckLeadingZero") = CheckLeadingZero
		ini.StringValue("DiscogsAutoTagWeb","ReleaseTag") = ReleaseTag
		ini.StringValue("DiscogsAutoTagWeb","CatalogTag") = CatalogTag
		ini.StringValue("DiscogsAutoTagWeb","CountryTag") = CountryTag
		ini.StringValue("DiscogsAutoTagWeb","FormatTag") = FormatTag
		ini.BoolValue("DiscogsAutoTagWeb","CheckVarious") = CheckVarious
		ini.StringValue("DiscogsAutoTagWeb","TxtVarious") = TxtVarious
		ini.BoolValue("DiscogsAutoTagWeb","CheckTitleFeaturing") = CheckTitleFeaturing
		ini.StringValue("DiscogsAutoTagWeb","TxtFeaturingName") = TxtFeaturingName
		ini.BoolValue("DiscogsAutoTagWeb","CheckFeaturingName") = CheckFeaturingName
		ini.BoolValue("DiscogsAutoTagWeb","CheckComment") = CheckComment
		ini.BoolValue("DiscogsAutoTagWeb","CheckUnselectNoTrackPos") = CheckUnselectNoTrackPos
		ini.BoolValue("DiscogsAutoTagWeb","SubTrackNameSelection") = SubTrackNameSelection
		ini.BoolValue("DiscogsAutoTagWeb","SkipNotChangedReleases") = SkipNotChangedReleases
		ini.BoolValue("DiscogsAutoTagWeb","ProcessOnlyDiscogs") = ProcessOnlyDiscogs
		ini.BoolValue("DiscogsAutoTagWeb","ProcessNoDiscogs") = ProcessNoDiscogs
		
		tmp = CountryFilterList.Item(0)
		For a = 1 To CountryList.Count - 1
			tmp = tmp & "," & CountryFilterList.Item(a)
		Next
		ini.StringValue("DiscogsAutoTagWeb","CurrentCountryFilter") = tmp
		tmp = MediaTypeFilterList.Item(0)
		For a = 1 To MediaTypeList.Count - 1
			tmp = tmp & "," & MediaTypeFilterList.Item(a)
		Next
		ini.StringValue("DiscogsAutoTagWeb","CurrentMediaTypeFilter") = tmp
		tmp = MediaFormatFilterList.Item(0)
		For a = 1 To MediaFormatList.Count - 1
			tmp = tmp & "," & MediaFormatFilterList.Item(a)
		Next
		ini.StringValue("DiscogsAutoTagWeb","CurrentMediaFormatFilter") = tmp
		tmp = YearFilterList.Item(0)
		For a = 1 To YearList.Count - 1
			tmp = tmp & "," & YearFilterList.Item(a)
		Next
		ini.StringValue("DiscogsAutoTagWeb","CurrentYearFilter") = tmp
	End If
	WriteLog "Stop SaveOptions"
	
End Sub

' Format Error Message
Sub FormatErrorMessage(ErrorMessage)

	WriteLog("Start FormatErrorMessage")
	WriteLog("ErrorMessage = " & ErrorMessage)
	Dim templateHTML, listBox, templateHTMLDoc, submitButton
	templateHTML = GetHeader()
	templateHTML = templateHTML &  "<tr>"
	templateHTML = templateHTML &  "<td colspan=4 align=center><p><b>" & ErrorMessage & "</b></p></td>"
	templateHTML = templateHTML &  "</tr>"
	templateHTML = templateHTML &  GetFooter()
	
	Set WebBrowser = SDB.Objects("WebBrowser")
	WebBrowser.SetHTMLDocument templateHTML
	Set WebBrowser2 = SDB.Objects("WebBrowser2")
	WebBrowser2.SetHTMLDocument ""
	
	Set templateHTMLDoc = WebBrowser.Interf.Document
	
	Set listBox = templateHTMLDoc.getElementById("alternative")
	Script.RegisterEvent listBox, "onchange", "Alternative"
	
	Set listBox = templateHTMLDoc.getElementById("filtermediatype")
	Script.RegisterEvent listBox, "onchange", "Filter"
	Set listBox = templateHTMLDoc.getElementById("filtermediaformat")
	Script.RegisterEvent listBox, "onchange", "Filter"
	Set listBox = templateHTMLDoc.getElementById("filtercountry")
	Script.RegisterEvent listBox, "onchange", "Filter"
	Set listBox = templateHTMLDoc.getElementById("filteryear")
	Script.RegisterEvent listBox, "onchange", "Filter"
	Set listBox = templateHTMLDoc.getElementById("load")
	Script.RegisterEvent listBox, "onchange", "Filter"
	Set submitButton = templateHTMLDoc.getElementById("showcountryfilter")
	Script.RegisterEvent submitButton, "onclick", "ShowCountryFilter"
	Set submitButton = templateHTMLDoc.getElementById("showmediatypefilter")
	Script.RegisterEvent submitButton, "onclick", "ShowMediaTypeFilter"
	Set submitButton = templateHTMLDoc.getElementById("showmediaformatfilter")
	Script.RegisterEvent submitButton, "onclick", "ShowMediaFormatFilter"
	Set submitButton = templateHTMLDoc.getElementById("showyearfilter")
	Script.RegisterEvent submitButton, "onclick", "ShowYearFilter"
	WriteLog("Stop FormatErrorMessage")
	
End Sub


Function URLEncodeUTF8(ByRef input)
	
	' urlencode a string with UTF8 encoding - yes, it is cryptic but it works!
	Dim i, result, CurrentChar
	Dim FirstByte, SecondByte, ThirdByte
	
	result = ""
	For i = 1 To Len(input)
		CurrentChar = Mid(input, i, 1)
		CurrentChar = AscW(CurrentChar)
		
		If (CurrentChar < 0) Then
			CurrentChar = CurrentChar + 65536
		End If
		
		If (CurrentChar >= 0) And (CurrentChar < 128) Then
			' 1 byte
			If(CurrentChar = 32) Then
				' replace space with "+"
				result = result & "+"
			Else
				' replace punctuation chars with "%hex"
				result = result & Escape(Chr(CurrentChar))
			End If
		End If
		
		If (CurrentChar >= 128) And (CurrentChar < 2048) Then
			' 2 bytes
			FirstByte  = &HC0 Xor ((CurrentChar And &HFFFFFFC0) \ &H40&)
			SecondByte = &H80 Xor (CurrentChar And &H3F)
			result = result & "%" & Hex(FirstByte) & "%" & Hex(SecondByte)
		End If
		
		If (CurrentChar >= 2048) And (CurrentChar < 65536) Then
			' 3 bytes
			FirstByte  = &HE0 Xor (((CurrentChar And &HFFFFF000) \ &H1000&) And &HF)
			SecondByte = &H80 Xor (((CurrentChar And &HFFFFFFC0) \ &H40&) And &H3F)
			ThirdByte  = &H80 Xor (CurrentChar And &H3F)
			result = result & "%" & Hex(FirstByte) & "%" & Hex(SecondByte) & "%" & Hex(ThirdByte)
		End If
	Next
	URLEncodeUTF8 = result
	
End Function


Function DecodeHtmlChars(Text)
	
	DecodeHtmlChars = Text
	DecodeHtmlChars = Replace(DecodeHtmlChars,"&quot;",	"""")
	DecodeHtmlChars = Replace(DecodeHtmlChars,"&lt;",	"<")
	DecodeHtmlChars = Replace(DecodeHtmlChars,"&gt;",	">")
	DecodeHtmlChars = Replace(DecodeHtmlChars,"&amp;",	"&")
	
End Function


Function EncodeHtmlChars(Text)
	
	EncodeHtmlChars= Text
	EncodeHtmlChars= Replace(EncodeHtmlChars, "&",	"&amp;")
	EncodeHtmlChars= Replace(EncodeHtmlChars,"""",	"&quot;")
	EncodeHtmlChars= Replace(EncodeHtmlChars,"<",	"&lt;")
	EncodeHtmlChars= Replace(EncodeHtmlChars, ">",	"&gt;")
	
End Function


Function CleanSearchString(Text)
	
	CleanSearchString = Text
	CleanSearchString = Replace(CleanSearchString,")", " ") 'remove paranthesis to avoid search errors (discogs bug)
	CleanSearchString = Replace(CleanSearchString,"(", " ") 'also clean other unneccessary characters
	CleanSearchString = Replace(CleanSearchString,"[", " ")
	CleanSearchString = Replace(CleanSearchString,"]", " ")
	CleanSearchString = Replace(CleanSearchString,".", " ")
	CleanSearchString = Replace(CleanSearchString,"@", " ")
	CleanSearchString = Replace(CleanSearchString,"_", " ")
	CleanSearchString = Replace(CleanSearchString,"?", " ")
	
End Function


Function CleanArtistName(artistname)
	
	CleanArtistName = DecodeHtmlChars(artistname)
	If InStr(CleanArtistName, " (") > 0 Then CleanArtistName = Left(CleanArtistName, InStrRev(CleanArtistName, " (") - 1)
	If InStr(CleanArtistName, ", The") > 0 Then CleanArtistName = "The " & Left(CleanArtistName, InStrRev(CleanArtistName, ", The") - 1)
	
End Function


Function AddAlternative(Alternative)
	
	Dim i
	If Trim(Alternative) <> "" Then
		For i = 0 To AlternativeList.Count - 1
			If AlternativeList.Item(i) = Trim(Alternative) Then
				Exit Function
			End If
		Next
		AlternativeList.Add Trim(Alternative)
	End If
	
End Function


Function AddAlternatives(Song)
	
	Dim SavedArtist, SavedTitle, SavedAlbum, SavedAlbumArtist, SavedFolderName, SavedFileName, Custom
	SavedArtist = Song.ArtistName
	SavedTitle = Song.Title
	SavedAlbum = Song.AlbumName
	SavedAlbumArtist = Song.AlbumArtistName
	SavedFolderName = Mid(Song.Path, 1, InStrRev(Song.Path,"\")-1)
	SavedFolderName = Mid(SavedFolderName, InStrRev(SavedFolderName,"\")+1)
	SavedFileName = Mid(Song.Path, 1, InStrRev(Song.Path,".")-1)
	SavedFileName = Mid(SavedFileName, InStrRev(SavedFileName,"\")+1)
	
	AddAlternative SavedFolderName
	If(InStr(SavedFolderName,"(") > 0) Then
		Custom = Mid(SavedFolderName,1,InStr(SavedFolderName,"(")-1)
		AddAlternative Custom
	End If
	If(InStr(SavedFolderName,"[") > 0) Then
		Custom = Mid(SavedFolderName,1,InStr(SavedFolderName,"[")-1)
		AddAlternative Custom
	End If
	AddAlternative SavedFileName
	If(InStr(SavedFileName,"(") > 0) Then
		Custom = Mid(SavedFileName,1,InStr(SavedFileName,"(")-1)
		AddAlternative Custom
	End If
	If(InStr(SavedFileName,"[") > 0) Then
		Custom = Mid(SavedFileName,1,InStr(SavedFileName,"[")-1)
		AddAlternative Custom
	End If
	AddAlternative Custom
	AddAlternative SavedArtist
	AddAlternative SavedTitle
	AddAlternative SavedAlbum
	AddAlternative SavedAlbumArtist
	If(InStr(SavedTitle,"(") > 0) Then
		Custom = Mid(SavedTitle,1,InStr(SavedTitle,"(")-1)
		AddAlternative Custom
	End If
	If(InStr(SavedTitle,"[") > 0) Then
		Custom = Mid(SavedTitle,1,InStr(SavedTitle,"[")-1)
		AddAlternative Custom
	End If
	AddAlternative SavedArtist & " " & SavedAlbum
	AddAlternative SavedAlbumArtist & " " & SavedAlbum
	AddAlternative SavedArtist & " " & SavedTitle
	AddAlternative SavedAlbumArtist & " " & SavedTitle
	
End Function


Function IsInteger(Str)

	Dim i, d
	IsInteger = True
	For i = 1 To Len(str)
		d = Mid(str, i, 1)
		If Asc(d) < 48 Or Asc(d) > 57 Then
			IsInteger = False
			Exit For
		End If
	Next
	
End Function


Function PackSpaces(Text)
	
	PackSpaces = Text
	PackSpaces = Replace(PackSpaces,"  ", " ") 'pack spaces
	PackSpaces = Replace(PackSpaces,"  ", " ") 'pack spaces left
	
End Function


Function search_involved(Text, SearchText)

	Dim i
	For i = 1 To UBound(Text)
		If Left(Text(i), Len(SearchText)) = SearchText Then
			search_involved = i
			Exit Function
		End If
	Next
	search_involved = -1
	
End Function


Function JSONParser_find_result(searchURL, ArrayName)
	
	Dim oXMLHTTP, r, f, a
	' use json api with vbsjson class at start of file now
	WriteLog "Start JSONParser_find_result"
	Set oXMLHTTP = CreateObject("MSXML2.XMLHTTP.6.0")
	
	Dim json
	Set json = New VbsJson
	
	Dim response
	Dim format, title, country, v_year, label, artist, Rtype, catNo, main_release, tmp, ReleaseDesc, FilterFound, SongCount, SongCountMax
	
	WriteLog "SearchURL=" & SearchURL
	Call oXMLHTTP.open("GET", searchURL, False)
	Call oXMLHTTP.setRequestHeader("Content-Type","application/json")
	Call oXMLHTTP.setRequestHeader("User-Agent","MediaMonkeyDiscogsAutoTagBatch/2.0 +http://mediamonkey.com")
	Call oXMLHTTP.send()
	
	If oXMLHTTP.Status = 200 Then
		Set response = json.Decode(oXMLHTTP.responseText)
		'check if any results
		'and add titles to drop down
		'msgbox response(ArrayName)(0)("title")
		
		SongCount = 1
		SongCountMax = response("pagination")("items")
		WriteLog ("SongCountMax=" & SongCountMax)
		
		For Each r In response(ArrayName)
			format = ""
			title = ""
			country = ""
			v_year = ""
			artist = ""
			label = ""
			Rtype = ""
			catNo = ""
			main_release = ""
			
			title = response(ArrayName)(r)("title")
			Set tmp = response(ArrayName)(r)
			If tmp.Exists("artist") Then
				artist = tmp("artist")
			End If
			If tmp.Exists("main_release") Then
				main_release = tmp("main_release")
			End If
			If ArrayName = "results" Then
				For Each f In response(ArrayName)(r)("format")
					format = format & response(ArrayName)(r)("format")(f) & ", "
				Next
				If Len(format) <> 0 Then format = Left(format, Len(format)-2)
			Else
				format = response(ArrayName)(r)("format")
			End If
			
			country = response(ArrayName)(r)("country")
			If ArrayName = "versions" Then
				If tmp.Exists("released") Then
					v_year = response(ArrayName)(r)("released")
				End If
			Else
				If tmp.Exists("year") Then
					v_year = response(ArrayName)(r)("year")
				End If
			End If
			If tmp.Exists("catno") Then
				catNo = response(ArrayName)(r)("catno")
			End If
			If tmp.Exists("type") Then
				Rtype = response(ArrayName)(r)("type")
			End If
			If ArrayName = "results" Then
				For Each f In response(ArrayName)(r)("label")
					If label <> "" Then
						If Left(label, Len(label)-2) <> response(ArrayName)(r)("label")(f) Then
							label = label & response(ArrayName)(r)("label")(f) & ", "
						End If
					Else
						label = response(ArrayName)(r)("label")(f) & ", "
					End If
				Next
				If Len(label) <> 0 Then label = Left(label, Len(label)-2)
			Else
				label = response(ArrayName)(r)("label")
			End If
			ReleaseDesc = ""
			Do
				If FilterMediaType = "Use MediaType Filter" And Format <> "" Then
					FilterFound = False
					For a = 1 To MediaTypeList.Count - 1
						If InStr(Format, MediaTypeList.Item(a)) <> 0 And MediaTypeFilterList.Item(a) = "1" Then FilterFound = True
					Next
					If FilterFound = False Then Exit Do
				End If
				If(FilterMediaType <> "None" And FilterMediaType <> "Use MediaType Filter" And InStr(format, FilterMediaType) = 0 And format <> "") Then Exit Do
				
				If FilterMediaFormat = "Use MediaFormat Filter" And format <> "" Then
					FilterFound = False
					For a = 1 To MediaFormatList.Count - 1
						If InStr(format, MediaFormatList.Item(a)) <> 0 And MediaFormatFilterList.Item(a) = "1" Then FilterFound = True
					Next
					If FilterFound = False Then Exit Do
				End If
				If(FilterMediaFormat <> "None" And FilterMediaFormat <> "Use MediaFormat Filter" And InStr(format, FilterMediaFormat) = 0 And Format <> "") Then Exit Do
				
				If FilterCountry = "Use Country Filter" And country <> "" Then
					FilterFound = False
					For a = 1 To CountryList.Count - 1
						If InStr(country, CountryList.Item(a)) <> 0 And CountryFilterList.Item(a) = "1" Then FilterFound = True
					Next
					If FilterFound = False Then Exit Do
				End If
				If(FilterCountry <> "None" And FilterCountry <> "Use Country Filter" And InStr(country, FilterCountry) = 0 And country <> "") Then Exit Do
				
				If FilterYear = "Use Year Filter" And v_year <> "" Then
					FilterFound = False
					For a = 1 To YearList.Count - 1
						If InStr(v_year, YearList.Item(a)) <> 0 And YearFilterList.Item(a) = "1" Then FilterFound = True
					Next
					If FilterFound = False Then Exit Do
				End If
				If(FilterYear <> "None" And FilterYear <> "Use Year Filter" And InStr(v_year, FilterYear) = 0 And v_year <> "") Then Exit Do
				
				If artist <> "" Then ReleaseDesc = ReleaseDesc & " " & artist End If
				If artist <> "" And title <> "" Then ReleaseDesc = ReleaseDesc & " -" End If
				If title <> "" Then ReleaseDesc = ReleaseDesc & " " & title End If
				If Format <> "" Then ReleaseDesc = ReleaseDesc & " [" & Format & "]" End If
				If Label <> "" Then ReleaseDesc = ReleaseDesc & " " & Label End If
				If Country <> "" Then ReleaseDesc = ReleaseDesc & " / " & Country End If
				If v_year <> "" Then ReleaseDesc = ReleaseDesc & " (" & v_year & ")" End If
				If catNo <> "" Then ReleaseDesc = ReleaseDesc & " catNo:" & catNo End If
				If Rtype = "master" Then ReleaseDesc = ReleaseDesc & " *" End If
				
				Combo.AddItem "(" & SongCount & "/" & SongCountMax & ") " & ReleaseDesc
				ResultsReleaseID.Add response(ArrayName)(r)("id")
				SongCount = SongCount + 1
			Loop While False
		Next
	End If
	WriteLog "Stop JSONParser_find_result"
	
End Function


Function ReloadMaster(SavedMasterId)
	
	Dim oXMLHTTP, masterURL
	WriteLog "Start ReloadMaster"
	masterURL = "http://api.discogs.com/masters/" & SavedMasterId
	Set oXMLHTTP = CreateObject("MSXML2.XMLHTTP.6.0")
	
	Dim json
	Set json = New VbsJson
	Dim response
	
	Call oXMLHTTP.open("GET", masterURL, False)
	Call oXMLHTTP.setRequestHeader("Content-Type","application/json")
	Call oXMLHTTP.setRequestHeader("User-Agent","MediaMonkeyDiscogsAutoTagBatch/2.0 +http://mediamonkey.com")
	Call oXMLHTTP.send()
	
	If oXMLHTTP.Status = 200 Then
		Set response = json.Decode(oXMLHTTP.responseText)
		If response.Exists("year") Then
			OriginalDate = response("year")
		Else
			OriginalDate = ""
		End If
	End If
	
	ReloadMaster = OriginalDate
	WriteLog "Stop ReloadMaster"
	
End Function


Function exchange_roman_numbers(Text)
	
	If Text = "I" Then Text = 1
	If Text = "II" Then Text = 2
	If Text = "III" Then Text = 3
	If Text = "IV" Then Text = 4
	If Text = "V" Then Text = 5
	If Text = "VI" Then Text = 6
	If Text = "VII" Then Text = 7
	If Text = "VIII" Then Text = 8
	If Text = "IX" Then Text = 9
	If Text = "X" Then Text = 10
	If Text = "XI" Then Text = 11
	If Text = "XII" Then Text = 12
	If Text = "XIII" Then Text = 13
	If Text = "XIV" Then Text = 14
	If Text = "XV" Then Text = 15
	If Text = "XVI" Then Text = 16
	If Text = "XVII" Then Text = 17
	If Text = "XVIII" Then Text = 18
	If Text = "XIX" Then Text = 19
	If Text = "XX" Then Text = 20
	exchange_roman_numbers = Text
	
End Function


Function get_release_ID(FirstTrack)
	
	CurrentReleaseID = ""
	If ReleaseTag = "Custom1" Then CurrentReleaseID = FirstTrack.Custom1
	If ReleaseTag = "Custom2" Then CurrentReleaseID = FirstTrack.Custom2
	If ReleaseTag = "Custom3" Then CurrentReleaseID = FirstTrack.Custom3
	If ReleaseTag = "Custom4" Then CurrentReleaseID = FirstTrack.Custom4
	If ReleaseTag = "Custom5" Then CurrentReleaseID = FirstTrack.Custom5
	If ReleaseTag = "Grouping" Then CurrentReleaseID = FirstTrack.Grouping
	If ReleaseTag = "ISRC" Then CurrentReleaseID = FirstTrack.ISRC
	If ReleaseTag = "Encoding" Then CurrentReleaseID = FirstTrack.Encoding
	If ReleaseTag = "Copyright" Then CurrentReleaseID = FirstTrack.Copyright
	
	WriteLog("CurrentReleaseID = " & CurrentReleaseID)
	get_release_ID = CurrentReleaseID
	
End Function


Sub WriteLog(Text)
	
	Dim filesys, filetxt, logdatei, tmpText, i
	'Const ForReading = 1, ForWriting = 2, ForAppending = 8
	logdatei = SDB.ScriptsPath & "Discogs_Batch_Script.log"
	Set filesys = CreateObject("Scripting.FileSystemObject")
	Set filetxt = filesys.OpenTextFile(logdatei, 8, True)
	If Left(Text, 4) = "Stop" Then
		cTab = cTab - 1
	End If
	tmpText = Time
	For i = 1 To cTab
		tmpText = tmpText & Chr(9)
	Next
	tmpText = tmpText & SDB.ToAscii(Text)
	If Left(Text, 5) = "Start" Then
		cTab = cTab + 1
	End If
	filetxt.WriteLine(tmpText)
	filetxt.Close
	
End Sub


Sub WriteLogInit
	
	Dim logdatei
	logdatei = SDB.ScriptsPath & "Discogs_Batch_Script.log"
	If SDB.Tools.FileSystem.FileExists(logdatei) = True Then
		SDB.Tools.FileSystem.DeleteFile(logdatei)
	End If
	
End Sub


Function AddToField(ByRef field, ByVal ftext)
	
	' for adding data to multi-valued fields
	If field = "" Then
		field = ftext
	Else
		field = field & Separator & ftext
	End If
	
End Function


Function LookForFeaturing(Text)

	Dim i, tmp, x
	tmp = Split(FeaturingKeywords, ",")
	For each x in tmp
		If LCase(Text) = LCase(x) Then
			LookForFeaturing = true
			Exit Function
		End If
	Next
	LookForFeaturing = false

End Function


Function CheckLeadingZeroTrackPosition(TrackPosition)
	
	Dim tmpSplit, tmpTrack
	If InStr(TrackPosition, "-") <> 0 Then
		tmpSplit = Split(TrackPosition, "-")
		TrackPosition = tmpSplit(1)
	End If
	If InStr(TrackPosition, ".") <> 0 Then
		tmpSplit = Split(TrackPosition, ".")
		TrackPosition = tmpSplit(1)
	End If
	If Left(TrackPosition, 1) = "0" Then
		CheckLeadingZeroTrackPosition = True
	Else
		CheckLeadingZeroTrackPosition = False
	End If
	
End Function


Class VbsJson
	'Author: Demon
	'Date: 2012/5/3
	'Website: http://demon.tw
	Private Whitespace, NumberRegex, StringChunk
	Private b, f, r, n, t
	
	Private Sub Class_Initialize
		Whitespace = " " & vbTab & vbCr & vbLf
		b = ChrW(8)
		f = vbFormFeed
		r = vbCr
		n = vbLf
		t = vbTab
		
		Set NumberRegex = New RegExp
		NumberRegex.Pattern = "(-?(?:0|[1-9]\d*))(\.\d+)?([eE][-+]?\d+)?"
		NumberRegex.Global = False
		NumberRegex.MultiLine = True
		NumberRegex.IgnoreCase = True
		
		Set StringChunk = New RegExp
		StringChunk.Pattern = "([\s\S]*?)([""\\\x00-\x1f])"
		StringChunk.Global = False
		StringChunk.MultiLine = True
		StringChunk.IgnoreCase = True
	End Sub
	
	'Return a JSON string representation of a VBScript data structure
	'Supports the following objects and types
	'+-------------------+---------------+
	'| VBScript          | JSON          |
	'+===================+===============+
	'| Dictionary        | object        |
	'+-------------------+---------------+
	'| Array             | array         |
	'+-------------------+---------------+
	'| String            | string        |
	'+-------------------+---------------+
	'| Number            | number        |
	'+-------------------+---------------+
	'| True              | true          |
	'+-------------------+---------------+
	'| False             | false         |
	'+-------------------+---------------+
	'| Null              | null          |
	'+-------------------+---------------+
	Public Function Encode(ByRef obj)
		Dim buf, i, c, g
		Set buf = CreateObject("Scripting.Dictionary")
		Select Case VarType(obj)
			Case vbNull
			buf.Add buf.Count, "null"
			Case vbBoolean
			If obj Then
				buf.Add buf.Count, "true"
			Else
				buf.Add buf.Count, "false"
			End If
			Case vbInteger, vbLong, vbSingle, vbDouble
			buf.Add buf.Count, obj
			Case vbString
			buf.Add buf.Count, """"
			For i = 1 To Len(obj)
				c = Mid(obj, i, 1)
				Select Case c
					Case """" buf.Add buf.Count, "\"""
					Case "\"  buf.Add buf.Count, "\\"
					Case "/"  buf.Add buf.Count, "/"
					Case b    buf.Add buf.Count, "\b"
					Case f    buf.Add buf.Count, "\f"
					Case r    buf.Add buf.Count, "\r"
					Case n    buf.Add buf.Count, "\n"
					Case t    buf.Add buf.Count, "\t"
					Case Else
					If AscW(c) >= 0 And AscW(c) <= 31 Then
						c = Right("0" & Hex(AscW(c)), 2)
						buf.Add buf.Count, "\u00" & c
					Else
						buf.Add buf.Count, c
					End If
				End Select
			Next
			buf.Add buf.Count, """"
			Case vbArray + vbVariant
			g = True
			buf.Add buf.Count, "["
			For Each i In obj
				If g Then g = False Else buf.Add buf.Count, ","
				buf.Add buf.Count, Encode(i)
			Next
			buf.Add buf.Count, "]"
			Case vbObject
			If TypeName(obj) = "Dictionary" Then
				g = True
				buf.Add buf.Count, "{"
				For Each i In obj
					If g Then g = False Else buf.Add buf.Count, ","
					buf.Add buf.Count, """" & i & """" & ":" & Encode(obj(i))
				Next
				buf.Add buf.Count, "}"
			Else
				Err.Raise 8732,,"None dictionary object"
			End If
			Case Else
			buf.Add buf.Count, """" & CStr(obj) & """"
		End Select
		Encode = Join(buf.Items, "")
	End Function
	
	'Return the VBScript representation of ``str(``
	'Performs the following translations in decoding
	'+---------------+-------------------+
	'| JSON          | VBScript          |
	'+===============+===================+
	'| object        | Dictionary        |
	'+---------------+-------------------+
	'| array         | Array             |
	'+---------------+-------------------+
	'| string        | String            |
	'+---------------+-------------------+
	'| number        | Double            |
	'+---------------+-------------------+
	'| true          | True              |
	'+---------------+-------------------+
	'| false         | False             |
	'+---------------+-------------------+
	'| null          | Null              |
	'+---------------+-------------------+
	Public Function Decode(ByRef str)
		'return base object
		Set Decode = ParseObject(str, 1)
	End Function
	
	Private Function ParseValue(ByRef str, ByRef idx)
		Dim c, ms
		
		idx = NextToken(str, idx)
		c = Mid(str, idx, 1)
		
		If c = "{" Then
			Set ParseValue = ParseObject(str, idx)
			Exit Function
		ElseIf c = "[" Then
			Set ParseValue = ParseArray(str, idx)
			Exit Function
		ElseIf c = """" Then
			idx = idx + 1
			ParseValue = ParseString(str, idx)
			Exit Function
		ElseIf c = "n" And StrComp("null", Mid(str, idx, 4)) = 0 Then
			idx = idx + 4
			ParseValue = Null
			Exit Function
		ElseIf c = "t" And StrComp("true", Mid(str, idx, 4)) = 0 Then
			idx = idx + 4
			ParseValue = True
			Exit Function
		ElseIf c = "f" And StrComp("false", Mid(str, idx, 5)) = 0 Then
			idx = idx + 5
			ParseValue = False
			Exit Function
		Else
			Set ms = NumberRegex.Execute(Mid(str, idx))
			If ms.Count = 1 Then
				idx = idx + ms(0).Length
				SetLocale "en-US"
				ParseValue = CDbl(ms(0))
				SetLocale 0
				Exit Function
			End If
		End If
		
		Err.Raise 8732,,"No JSON object could be ParseValued"
	End Function
	
	Private Function ParseObject(ByRef str, ByRef idx)
		Dim c, key, value
		Set ParseObject = CreateObject("Scripting.Dictionary")
		
		idx = NextToken(str, idx)
		
		c = Mid(str, idx, 1)
		
		If c = "{" Then
			idx = NextToken(str,idx+1)
		Else
			Err.Raise 8732,,"Expected { to begin Object"
		End If
		
		c = Mid(str, idx, 1)
		
		Do
			If c <> """" And c <> "}" Then
				
				Err.Raise 8732,,"Expecting property name or } near: " & Mid(str,idx)
				
			ElseIf c = """" Then
				
				idx = idx + 1
				key = ParseString(str, idx)
				
				idx = NextToken(str, idx)
				If Mid(str, idx, 1) <> ":" Then
					Err.Raise 8732,,"Expecting : delimiter near: " & Mid(str,idx)
				End If
				
				' skip : and whitespace after key
				idx = NextToken(str, idx + 1)
				
				' check for object or array value
				If Mid(str,idx,1) = "{" Or Mid(str,idx,1) = "[" Then
					Set value = ParseValue(str, idx)
				Else
					value = ParseValue(str,idx)
				End If
				
				ParseObject.Add key, value
			End If
			
			c = Mid(str,idx,1)
			
			If c = "}" Then
				idx = NextToken(str,idx+1)
				Exit Function
			End If
			
			If c <> "," Then
				
				Err.Raise 8732,,"Expecting , delimiter near: " & Mid(str,idx)
				
			End If
			
			'skip , and whitespace after value
			idx = NextToken(str, idx + 1)
			c = Mid(str, idx, 1)
			If c <> """" Then
				Err.Raise 8732,,"Expecting property name"
			End If
		Loop
	End Function
	
	Private Function ParseArray(ByRef str, ByRef idx)
		Dim c, values, value
		Set ParseArray = CreateObject("Scripting.Dictionary")
		
		idx = NextToken(str, idx)
		c = Mid(str, idx, 1)
		
		If c = "[" Then
			idx = NextToken(str,idx+1)
		Else
			Err.Raise 8732,,"Expected [ to begin Array"
		End If
		
		Do
			c = Mid(str, idx, 1)
			
			If c = "]" Then
				idx = NextToken(str,idx+1)
				Exit Function
			End If
			
			ParseArray.Add ParseArray.Count, ParseValue(str, idx)
			
			c = Mid(str, idx, 1)
			
			If c = "]" Then
				idx = NextToken(str, idx+1)
				Exit Function
			End If
			
			If c <> "," Then
				Err.Raise 8732,,"Expecting , delimiter near: " & Mid(str,idx)
			End If
			
			idx = NextToken(str,idx+1)
			
		Loop
	End Function
	
	Private Function ParseString(ByRef str, ByRef idx)
		Dim chunks, content, terminator, ms, esc, char
		Set chunks = CreateObject("Scripting.Dictionary")
		
		Do
			Set ms = StringChunk.Execute(Mid(str, idx))
			If ms.Count = 0 Then
				Err.Raise 8732,,"Unterminated string starting"
			End If
			
			content = ms(0).Submatches(0)
			terminator = ms(0).Submatches(1)
			If Len(content) > 0 Then
				chunks.Add chunks.Count, content
			End If
			
			idx = idx + ms(0).Length
			
			If terminator = """" Then
				Exit Do
			ElseIf terminator <> "\" Then
				Err.Raise 8732,,"Invalid control character"
			End If
			
			esc = Mid(str, idx, 1)
			
			If esc <> "u" Then
				Select Case esc
					Case """" char = """"
					Case "\"  char = "\"
					Case "/"  char = "/"
					Case "b"  char = b
					Case "f"  char = f
					Case "n"  char = n
					Case "r"  char = r
					Case "t"  char = t
					Case Else Err.Raise 8732,,"Invalid escape"
				End Select
				idx = idx + 1
			Else
				char = ChrW("&H" & Mid(str, idx + 1, 4))
				idx = idx + 5
			End If
			
			chunks.Add chunks.Count, char
		Loop
		
		ParseString = Join(chunks.Items, "")
	End Function
	
	Private Function NextToken(ByRef str, ByVal idx)
		Do While idx <= Len(str) And InStr(Whitespace, Mid(str, idx, 1)) > 0
			idx = idx + 1
		Loop
		NextToken = idx
	End Function
	
End Class


Function searchKeyword(Keywords, Role, AlbumRole, artistName)
	
	Dim tmp, x
	tmp = Split(Keywords, ",")
	For Each x In tmp
		If LCase(Role) = LCase(x) Then
			If InStr(AlbumRole, artistName) = 0 Then
				If AlbumRole = "" Then
					AlbumRole = artistName
				Else
					AlbumRole = AlbumRole & Separator & artistName
				End If
				searchKeyword = AlbumRole
			Else
				searchKeyword = "ALREADY_INSIDE_ROLE"
			End If
			Exit For
		End If
	Next
	
End Function


Function getCustom(Tag)
	
	If Tag = "Custom1" Then getCustom = ini.StringValue("CustomFields","Fld1Name")
	If Tag = "Custom2" Then getCustom = ini.StringValue("CustomFields","Fld2Name")
	If Tag = "Custom3" Then getCustom = ini.StringValue("CustomFields","Fld3Name")
	If Tag = "Custom4" Then getCustom = ini.StringValue("CustomFields","Fld4Name")
	If Tag = "Custom5" Then getCustom = ini.StringValue("CustomFields","Fld5Name")
	
End Function


Sub NewSearch(CurrentSelectedAlbum)

	Dim i, AlbumID, Label, iter, objSongData, currentTrack
	Dim searchAlbum, searchArtist, searchTerm, itm2

	WriteLog "Start NewSearch"

	Set NewTrackList = SDB.NewSongList
	Set AlternativeList = SDB.NewStringList

	For i = 0 To 1000
		UnselectedTracks(i) = ""
	Next

	Dim QueryString

	SavedMasterId = ""
	SavedReleaseId = ""
	SavedSearchTerm = ""
	SavedArtistId = ""
	SavedLabelId = ""

	FirstTrack = ""
	AlbumArtURL = ""
	AlbumArtThumbNail = ""
	iMaxTracks = 0
	LastDisc = ""
	SelectAll = True
	CheckDiscNum = True
	CheckTrackNum = True
	ResultsReleaseID = ""
	CurrentReleaseID = ""
	QueryString = ""

	OriginalDate = ""
	ReleaseDate = ""

	RadioBoxCheck = -1

	WriteLog("AlbumIDList.Count=" & AlbumIDList.Count)
	WriteLog("CurrentSelectedAlbum=" & CurrentSelectedAlbum)
	If AlbumIDList.Count > CurrentSelectedAlbum Then

		AlbumID = AlbumIDList.Item(CurrentSelectedAlbum)

		QueryString = "IDAlbum = """ & AlbumID & """ ORDER BY CAST(Discnumber AS SIGNED INTEGER), CAST(Tracknumber AS SIGNED INTEGER)"
		Set iter = SDB.Database.QuerySongs(QueryString)
		Do While Not iter.EOF
			Set objSongData = iter.Item
			NewTrackList.Add(objSongData)
			iter.Next
		Loop
		Set iter=Nothing
		
		For i = 0 To NewTrackList.Count-1
			Set currentTrack = NewTrackList.Item(i)
			WriteLog("Song " & i & "  /  Artist=" & currentTrack.ArtistName & "  /  Title=" & currentTrack.Title & "  /  Album=" & currentTrack.AlbumName)
		Next
		
		SearchAlbum = currentTrack.AlbumName
		SearchArtist = currentTrack.AlbumArtistName
		SearchTerm = currentTrack.AlbumArtistName & " " & currentTrack.AlbumName
		WriteLog "NewSearch SearchArtist=" & SearchArtist
		WriteLog "NewSearch SearchAlbum=" & SearchAlbum
		
		WriteLog "Number of Query Songs=" & NewTrackList.Count
		
		AddAlternative SearchTerm
		AddAlternative SearchArtist
		AddAlternative SearchAlbum
		
		Set itm2 = NewTrackList.item(0)
		WriteLog "1.Songtitle=" & itm2.Title
		
		Dim AlbumArt
		Set AlbumArt = itm2.AlbumArt
		If AlbumArt.Count > 0 Then
			WriteLog AlbumArt.Count & " Cover-Images vorhanden"
			If CheckNotAlwaysSaveImage = True Then
				WriteLog "CheckNotAlwaysSaveImage = true -> Covers werden nicht gespeichert"
				CheckCover = False
				CheckSmallCover = False
			Else
				WriteLog "CheckNotAlwaysSaveImage = false -> Covers werden zum speichern vorgemerkt"
			End If
		End If
		
		For i = 0 To NewTrackList.Count - 1
			AddAlternatives NewTrackList.item(i)
		Next
		
		Rem FilterMediaType = "None"
		Rem FilterMediaFormat = "None"
		Rem FilterCountry = "None"
		Rem FilterYear = "None"
		CurrentLoadType = "Search Results"
		
		
		
		Dim FormWidth
		Dim FormHeight
		FormWidth = 900
		FormHeight = 600

		Dim SearchForm : Set SearchForm = UI.NewForm
		SearchForm.Common.SetRect 0, 0, FormWidth, FormHeight
		SearchForm.FormPosition = 4
		SearchForm.SavePositionName = "TheDiscogsWindow"
		SearchForm.Caption = "Discogs Search"
		'FormBorderStyle = 0
		SearchForm.Common.MinWidth = 800
		SearchForm.Common.MinHeight = 600
		SearchFormWidth = SearchForm.Common.Width
		SearchForm.StayOnTop = True
		SearchForm.Common.Visible = True                ' Only show the form, don't wait for user input
		SDB.Objects("SearchForm") = SearchForm  ' Save reference to the form somewhere, otherwise it would simply disappear
		
		Set Head = UI.NewPanel(SearchForm)
		Head.Common.Align = 1   ' Top
		Head.Common.Height = 30
		
		Dim Bottom : Set Bottom = UI.NewPanel(SearchForm)
		Bottom.Common.Align = 2   ' Bottom
		Bottom.Common.Height = 30
		
		Set Label = UI.NewLabel(Bottom)
		Label.Common.FontBold = True
		Label.Caption = "Processing " & CurrentSelectedAlbum+1 & " of " & AlbumIDList.Count & " Albums"
		Label.Common.SetRect 105, 5, 260, 25
		SDB.Objects("ProcessLabel") = Label
		
		Dim Btn10 : Set Btn10 = UI.NewButton(Bottom)
		Btn10.Common.ControlName = "Btn10"
		Btn10.Common.SetRect 5, 5, 35, 20
		Btn10.Caption = SDB.Localize("Up")
		Btn10.Common.Hint = "Move the selected track one position up"
		Btn10.Common.Anchors = 6
		Script.RegisterEvent Btn10, "OnClick", "Btn10Click"
		
		Dim Btn11 : Set Btn11 = UI.NewButton(Bottom)
		Btn11.Common.ControlName = "Btn11"
		Btn11.Common.SetRect 45, 5, 35, 20
		Btn11.Caption = SDB.Localize("Down")
		Btn11.Common.Hint = "Move the selected track one position down"
		Btn11.Common.Anchors = 6
		Script.RegisterEvent Btn11, "OnClick", "Btn11Click"
		
		Dim Combo : Set Combo = UI.NewDropDown(Head)
		Combo.Common.SetRect 5, 5, SearchFormWidth -550, 20
		Combo.Style = 2     ' List
		Script.RegisterEvent Combo, "OnSelect", "ComboChange"
		
		Dim Btn2 : Set Btn2 = UI.NewButton(Head)
		Btn2.Common.ControlName = "Btn2"
		Btn2.Common.SetRect SearchForm.Common.Width -310, 5, 80, 20
		Btn2.Caption = SDB.Localize("Update")
		Btn2.Common.Hint = "Update the tag(s) with different content"
		Btn2.Common.Anchors = 6
		Script.RegisterEvent Btn2, "OnClick", "Btn2Click"
		SDB.Objects("Btn2") = Btn2
		
		Dim Btn3 : Set Btn3 = UI.NewButton(Head)
		Btn3.Common.ControlName = "Btn3"
		Btn3.Common.SetRect SearchForm.Common.Width -210, 5, 80, 20
		Btn3.Caption = SDB.Localize("Skip")
		Btn3.Common.Hint = "Skip this album"
		Btn3.Common.Anchors = 6
		Script.RegisterEvent Btn3, "OnClick", "Btn3Click"
		
		Dim Btn4: Set Btn4 = UI.NewButton(Head)
		Btn4.Common.ControlName = "Btn4"
		Btn4.Common.SetRect SearchForm.Common.Width -110, 5, 80, 20
		Btn4.Caption = SDB.Localize("Cancel")
		Btn4.Common.Hint = "Stop and exit the script"
		Btn4.Common.Anchors = 6
		Script.RegisterEvent Btn4, "OnClick", "Btn4Click"
		
		Dim Btn5 : Set Btn5 = UI.NewButton(Head)
		Btn5.Common.ControlName = "Btn5"
		Btn5.Common.SetRect SearchForm.Common.Width -470, 5, 140, 20
		Btn5.Caption = "Update select. tracks"
		Btn5.Common.Hint = "Update the track table below, after changing the selected tracks"
		Btn5.Common.Anchors = 6
		Script.RegisterEvent Btn5, "OnClick", "Btn5Click"
		
		Set WebBrowser = UI.NewActiveX(SearchForm, "Shell.Explorer")
		WebBrowser.Common.Align = 0
		WebBrowser.Common.ControlName = "WebBrowser"
		WebBrowser.Common.Top = 30
		WebBrowser.Common.Left = 0
		WebBrowser.Common.Height = 600
		WebBrowser.Common.Width = 1650
		SDB.Objects("WebBrowser") = WebBrowser
		WebBrowser.Interf.Visible = True
		WebBrowser.Common.BringToFront
		
		Set WebBrowser2 = UI.NewActiveX(SearchForm, "Shell.Explorer")
		WebBrowser2.Common.Align = 0      ' Fill whole client rectangle
		WebBrowser2.Common.ControlName = "WebBrowser2"
		WebBrowser2.Common.Top = 630
		WebBrowser2.Common.Left = 0
		WebBrowser2.Common.Height = 300
		WebBrowser2.Common.Width = 1650
		SDB.Objects("WebBrowser2") = WebBrowser2
		WebBrowser2.Interf.Visible = True
		WebBrowser2.Common.BringToFront
		SDB.ProcessMessages

		If NewTrackList.Count > 0 Then
			Set FirstTrack = NewTrackList.item(0)
			SavedReleaseId = get_release_ID(FirstTrack)
			SavedSearchTerm = SearchTerm
		End If
		
		If (SavedReleaseId <> "" And ProcessOnlyDiscogs = True) Or (SavedReleaseId = "" And ProcessNoDiscogs = True) Or (ProcessOnlyDiscogs = False And ProcessNoDiscogs = False) Then
		
			FindResults SavedSearchTerm, SearchArtist, SearchAlbum
			
			If ErrorMessage <> "" Then
				FormatErrorMessage ErrorMessage
			Else
				WebBrowser.SetHTMLDocument templateHTML
				WebBrowser2.SetHTMLDocument tracklistHTML
				ShowResult 0
			End If
		Else
			If SavedReleaseId = "" And ProcessOnlyDiscogs = True Then
				WriteLog "Album have no ReleaseID"
				cReleasesOnlyDiscogsSkip = cReleasesOnlyDiscogsSkip + 1
			End If
			If SavedReleaseId <> "" And ProcessNoDiscogs = True Then
				WriteLog "Album have ReleaseID"
				cReleasesNoDiscogsSkip = cReleasesNoDiscogsSkip + 1
			End If
			CurrentSelectedAlbum = CurrentSelectedAlbum + 1
		End If
	End If
	WriteLog "Stop NewSearch"

End Sub


Sub TrackPosUp(SongNr)

	Dim g, tmpTrackList
	Dim templateHTMLDoc
	If SongNr = 0 Then Exit Sub
	Set tmpTrackList = SDB.NewSongList
	If SongNr = 1 Then
		tmpTrackList.add NewTrackList.Item(1)
		tmpTrackList.add NewTrackList.Item(0)
	End If
	If SongNr > 1 Then
		For g = 0 to SongNr-2
			tmpTrackList.add NewTrackList.Item(g)
		Next
		tmpTrackList.add NewTrackList.Item(SongNr)
		tmpTrackList.add NewTrackList.Item(SongNr-1)
	End If
	For g = SongNr+1 to NewTrackList.Count-1
		tmpTrackList.add NewTrackList.Item(g)
	Next
	Set NewTrackList = SDB.NewSongList
	For g = 0 to tmpTrackList.Count-1
		NewTrackList.add tmpTrackList.Item(g)
	Next
	RadioBoxCheck = SongNr - 1

End Sub


Sub TrackPosDown(SongNr)
	
	Dim g, tmpTrackList
	If SongNr = NewTrackList.count-1 Then Exit Sub
	Set tmpTrackList = SDB.NewSongList
	For g = 0 to SongNr-1
		tmpTrackList.add NewTrackList.Item(g)
	Next
	tmpTrackList.add NewTrackList.Item(SongNr+1)
	tmpTrackList.add NewTrackList.Item(SongNr)
	If SongNr+2 < NewTrackList.count-1 Then
		For g = SongNr+2 to NewTrackList.Count-1
			tmpTrackList.add NewTrackList.Item(g)
		Next
	End If
	Set NewTrackList = SDB.NewSongList
	For g = 0 to tmpTrackList.Count-1
		NewTrackList.add tmpTrackList.Item(g)
	Next
	RadioBoxCheck = SongNr + 1

End Sub


Sub ComboChange(Combo)
	
	WriteLog "Start ComboChange"
	Dim Index
	Index = Combo.ItemIndex
	WriteLog "ComboChange Index=" & Index
	RadioBoxCheck = -1
	ShowResult Index
	WriteLog "Stop ComboChange"
	
End Sub


Sub OnClose(Form)
	
	WriteLog "Sub OnClose"
	FinishSearch Form
	
End Sub


Sub Btn2Click
	
	WriteLog "update selected"
	UpdateAlbumTracks
	CurrentSelectedAlbum = CurrentSelectedAlbum + 1
	cReleasesUpdate = cReleasesUpdate + 1
	If AlbumIDList.Count <= CurrentSelectedAlbum Then
		FinishSearch Form
	Else
		NewSearch CurrentSelectedAlbum
	End If
	
End Sub


Sub Btn3Click
	
	WriteLog "skip selected"
	CurrentSelectedAlbum = CurrentSelectedAlbum + 1
	cReleasesSkip = cReleasesSkip + 1
	If AlbumIDList.Count <= CurrentSelectedAlbum Then
		FinishSearch Form
	Else
		NewSearch CurrentSelectedAlbum
	End If
	
End Sub


Sub Btn4Click
	
	WriteLog "cancel selected"
	FinishSearch Form
	
End Sub


Sub Btn5Click
	
	Dim templateHTMLDoc, i, checkBox
	
	Set WebBrowser = SDB.Objects("WebBrowser")
	Set templateHTMLDoc = WebBrowser.Interf.Document
	
	For i = 0 To iMaxTracks - 1
		Set checkBox = templateHTMLDoc.getElementById("unselected["&i&"]")
		If checkBox.Checked Then
			UnselectedTracks(i) = ""
		Else
			UnselectedTracks(i) = "x"
		End If
	Next
	
	ReloadResults
	
End Sub


Sub Btn10Click
	
	Dim g, Radiobox
	Dim templateHTMLDoc
	Set WebBrowser2 = SDB.Objects("WebBrowser2")
	Set templateHTMLDoc = WebBrowser2.Interf.Document
	For g = 0 To NewTrackList.Count-1
		Set RadioBox = templateHTMLDoc.getElementById(g)
		If RadioBox.checked Then
			TrackPosUp g
			ReloadResults
			Exit Sub
		End If
	Next
	
End Sub


Sub Btn11Click
	
	Dim g, Radiobox
	Dim templateHTMLDoc
	Set WebBrowser2 = SDB.Objects("WebBrowser2")
	Set templateHTMLDoc = WebBrowser2.Interf.Document
	For g = 0 To NewTrackList.Count-1
		Set RadioBox = templateHTMLDoc.getElementById(g)
		If RadioBox.checked Then
			TrackPosDown g
			ReloadResults
			Exit Sub
		End If
	Next
	
End Sub


Sub UpdateAlbumTracks()

	Dim i, checkBox, res, j
	Dim templateHTMLDoc
	WriteLog "Start UpdateAlbumTracks"

	Set WebBrowser = SDB.Objects("WebBrowser")
	Set templateHTMLDoc = WebBrowser.Interf.Document
	
	For i = 0 To iMaxTracks - 1
		Set checkBox = templateHTMLDoc.getElementById("unselected["&i&"]")
		If (checkBox.Checked And UnselectedTracks(i) = "x") Or (Not checkBox.Checked And UnselectedTracks(i) = "") Then
			res = SDB.MessageBox( "You don't have updated the track list after changing it. Should i update the track-list ?", mtConfirmation, Array(mbYes, mbNo, mbCancel))
			If res = 2 Then Exit Sub 'Cancel
			If res = 6 Then
				For j = 0 To iMaxTracks - 1
					Set checkBox = templateHTMLDoc.getElementById("unselected["&j&"]")
					If checkBox.Checked Then
						UnselectedTracks(j) = ""
					Else
						UnselectedTracks(j) = "x"
					End If
				Next
				Exit For
			End If
		End If
	Next
	
	
	WriteLog "Start CoverUpdate"
	If CheckCover Or CheckSmallCover Then
		Dim path : path = NewTrackList.Item(0).Path
		Dim k : k = InStrRev(path,"\")
		If k > 0 Then
			path = Left(path,k)
			path = path & "folder.jpg"
			If CheckCover Then
				SDB.Downloader.DownloadFile AlbumArtURL, path, False
			Else
				SDB.Downloader.DownloadFile AlbumArtThumbNail, path, False
			End If
		End If
	End If
	WriteLog "Stop CoverUpdate"
	
	j = 0
	
	For i = 0 To NewTrackList.Count - 1
		
		Do
			If j < tracks.count Then
				If UnselectedTracks(j) = "" Then
					
					If CheckArtist Then NewTrackList.Item(i).ArtistName = AlbumArtistTitle
					If CheckAlbumArtist Then NewTrackList.Item(i).AlbumArtistName = AlbumArtist
					If CheckAlbum Then NewTrackList.Item(i).AlbumName = AlbumTitle
					
					If CheckDate Then
						If Len(ReleaseDate) > 4 Then
							NewTrackList.Item(i).Year = Mid(ReleaseDate,7,4)
							NewTrackList.Item(i).Month = Mid(ReleaseDate,4,2)
							NewTrackList.Item(i).Day = Mid(ReleaseDate,1,2)
						ElseIf ReleaseDate = "" Then
							NewTrackList.Item(i).Year = -1
						ElseIf IsNumeric(ReleaseDate) Then
							NewTrackList.Item(i).Year = ReleaseDate
						End If
					End If
					
					If CheckOrigDate Then
						If Len(OriginalDate) > 4 Then
							NewTrackList.Item(i).OriginalYear = Mid(OriginalDate,7,4)
							NewTrackList.Item(i).OriginalMonth = Mid(OriginalDate,4,2)
							NewTrackList.Item(i).OriginalDay = Mid(OriginalDate,1,2)
						ElseIf OriginalDate = "" Then
							NewTrackList.Item(i).OriginalYear = -1
						ElseIf IsNumeric(OriginalDate) Then
							NewTrackList.Item(i).OriginalYear = OriginalDate
						End If
					End If
					
					If CheckCover Or CheckSmallCover Then
						Set Art = NewTrackList.Item(i).AlbumArt
						If Art.Count > 0 Then Art.Delete(0)
						Set img = Art.AddNew
						img.RelativePicturePath = "folder.jpg"
						img.Description = ""
						img.ItemType = 3
						img.ItemStorage = 1
						Art.UpdateDB
					End If
					
					If CheckStyleField = "Default (stored with Genre)" Then
						If CheckGenre And CheckStyle Then
							NewTrackList.Item(i).Genre = Genres & Separator & Styles
							If Genres = "" Then NewTrackList.Item(i).Genre = Styles
							If Styles = "" Then NewTrackList.Item(i).Genre = Genres
						ElseIf CheckGenre Then
							NewTrackList.Item(i).Genre = Genres
						ElseIf CheckStyle Then
							NewTrackList.Item(i).Genre = Styles
						End If
					Else
						If CheckGenre Then
							NewTrackList.Item(i).Genre = Genres
						End If
						If CheckStyle Then
							If CheckStyleField = "Custom1" Then NewTrackList.Item(i).Custom1 = Styles
							If CheckStyleField = "Custom2" Then NewTrackList.Item(i).Custom2 = Styles
							If CheckStyleField = "Custom3" Then NewTrackList.Item(i).Custom3 = Styles
							If CheckStyleField = "Custom4" Then NewTrackList.Item(i).Custom4 = Styles
							If CheckStyleField = "Custom5" Then NewTrackList.Item(i).Custom5 = Styles
						End If
					End If
					If CheckLabel Then NewTrackList.Item(i).Publisher = theLabels
					
					If CheckComment Then NewTrackList.Item(i).Comment = Comment
					
					If CheckRelease Then
						If ReleaseTag = "Custom1" Then NewTrackList.Item(i).Custom1 = CurrentReleaseID
						If ReleaseTag = "Custom2" Then NewTrackList.Item(i).Custom2 = CurrentReleaseID
						If ReleaseTag = "Custom3" Then NewTrackList.Item(i).Custom3 = CurrentReleaseID
						If ReleaseTag = "Custom4" Then NewTrackList.Item(i).Custom4 = CurrentReleaseID
						If ReleaseTag = "Custom5" Then NewTrackList.Item(i).Custom5 = CurrentReleaseID
						If ReleaseTag = "Grouping" Then NewTrackList.Item(i).Grouping = CurrentReleaseID
						If ReleaseTag = "ISRC" Then NewTrackList.Item(i).ISRC = CurrentReleaseID
						If ReleaseTag = "Encoding" Then NewTrackList.Item(i).Encodiung = CurrentReleaseID
						If ReleaseTag = "Copyright" Then NewTrackList.Item(i).Copyright = CurrentReleaseID
					End If
					
					If CheckCatalog Then
						If CatalogTag = "Custom1" Then NewTrackList.Item(i).Custom1 = theCatalogs
						If CatalogTag = "Custom2" Then NewTrackList.Item(i).Custom2 = theCatalogs
						If CatalogTag = "Custom3" Then NewTrackList.Item(i).Custom3 = theCatalogs
						If CatalogTag = "Custom4" Then NewTrackList.Item(i).Custom4 = theCatalogs
						If CatalogTag = "Custom5" Then NewTrackList.Item(i).Custom5 = theCatalogs
					End If
					
					If CheckCountry Then
						If CountryTag = "Custom1" Then NewTrackList.Item(i).Custom1 = theCountry
						If CountryTag = "Custom2" Then NewTrackList.Item(i).Custom2 = theCountry
						If CountryTag = "Custom3" Then NewTrackList.Item(i).Custom3 = theCountry
						If CountryTag = "Custom4" Then NewTrackList.Item(i).Custom4 = theCountry
						If CountryTag = "Custom5" Then NewTrackList.Item(i).Custom5 = theCountry
					End If
					
					If CheckFormat Then
						If FormatTag = "Custom1" Then NewTrackList.Item(i).Custom1 = theFormat
						If FormatTag = "Custom2" Then NewTrackList.Item(i).Custom2 = theFormat
						If FormatTag = "Custom3" Then NewTrackList.Item(i).Custom3 = theFormat
						If FormatTag = "Custom4" Then NewTrackList.Item(i).Custom4 = theFormat
						If FormatTag = "Custom5" Then NewTrackList.Item(i).Custom5 = theFormat
					End If
					
					If StrComp(NewTrackList.Item(i).Title, Tracks.Item(j),1) <> 0 Then NewTrackList.Item(i).Title = Tracks.Item(j)
					If CheckArtist Then NewTrackList.Item(i).ArtistName = ArtistTitles.Item(j)
					If CheckTrackNum Then NewTrackList.Item(i).TrackOrderStr = TracksNum.Item(j)
					If CheckDiscNum Then NewTrackList.Item(i).DiscNumberStr = TracksCD.Item(j)
					If CheckInvolved Then NewTrackList.Item(i).InvolvedPeople = InvolvedArtists.Item(j)
					If CheckLyricist Then NewTrackList.Item(i).Lyricist = Lyricists.Item(j)
					If CheckComposer Then NewTrackList.Item(i).Author = Composers.Item(j)
					If CheckConductor Then NewTrackList.Item(i).Conductor = Conductors.Item(j)
					If CheckProducer Then NewTrackList.Item(i).Producer = Producers.Item(j)
					j = j + 1
					Exit Do
				Else
					j = j + 1
				End If
			Else
				Exit Do
			End If
		Loop While True
	Next
	NewTrackList.UpdateAll
	WriteLog "Stop UpdateAlbumTracks"
	
End Sub


Function trackliste_aufbauen (Tracks, TracksNum, TracksCD, Durations, AlbumArtist, AlbumArtistTitle, ArtistTitles, AlbumTitle, ReleaseDate, OriginalDate, Genres, Styles, theLabels, theCountry, AlbumArtThumbNail, CurrentReleaseID, theCatalogs, Lyricists, Composers, Conductors, Producers, InvolvedArtists, theFormat, theMaster, Comment)
	
	WriteLog "Start Trackliste_aufbauen"
	Dim GenStyle
	Dim c, j
	
	tracklistHTML = "<HTML>"
	tracklistHTML = tracklistHTML & "<HEAD>"
	tracklistHTML = tracklistHTML & "<style type=""text/css"" media=""screen"">"
	tracklistHTML = tracklistHTML & ".tabletext { font-family: Arial, Helvetica, sans-serif; font-size: 8pt;}"
	tracklistHTML = tracklistHTML & "</style>"
	tracklistHTML = tracklistHTML & "</HEAD>"
	tracklistHTML = tracklistHTML & "<body bgcolor=""#FFFFFF"">"
	tracklistHTML = tracklistHTML & "<table border=1 cellspacing=2 cellpadding=5 class=tabletext>"
	tracklistHTML = tracklistHTML & "<tr>"
	
	'table head
	tracklistHTML = tracklistHTML & "<th><b>C</b></th>"
	tracklistHTML = tracklistHTML & "<th><b>#</b></th>"
	tracklistHTML = tracklistHTML & "<th><b>  " & SDB.Localize("Track #") & "  </b></th>"
	tracklistHTML = tracklistHTML & "<th><b>  " & SDB.Localize("Disc #") & "  </b></th>"
	tracklistHTML = tracklistHTML & "<th><b>" & SDB.Localize("Title") & "</b></th>"
	tracklistHTML = tracklistHTML & "<th><b>" & SDB.Localize("Artist(s)") & "</b></th>"
	tracklistHTML = tracklistHTML & "<th><b>" & SDB.Localize("Album") & "</b></th>"
	tracklistHTML = tracklistHTML & "<th><b>" & SDB.Localize("Album artist(s)") & "</b></th>"
	tracklistHTML = tracklistHTML & "<th><b>" & SDB.Localize("Genre(s)") & "</b></th>"
	tracklistHTML = tracklistHTML & "<th><b>" & SDB.Localize("Date") & "</b></th>"
	tracklistHTML = tracklistHTML & "<th><b>" & SDB.Localize("Original date") & "</b></th>"
	tracklistHTML = tracklistHTML & "<th><b>" & SDB.Localize("Filename") & "</b></th>"
	tracklistHTML = tracklistHTML & "<th><b>" & SDB.Localize("Composer(s)") & "</b></th>"
	tracklistHTML = tracklistHTML & "<th><b>" & SDB.Localize("Conductor(s)") & "</b></th>"
	tracklistHTML = tracklistHTML & "<th><b>" & SDB.Localize("Producer(s)") & "</b></th>"
	tracklistHTML = tracklistHTML & "<th><b>" & SDB.Localize("Lyricist(s)") & "</b></th>"
	tracklistHTML = tracklistHTML & "<th><b>" & SDB.Localize("Involved People") & "</b></th>"
	tracklistHTML = tracklistHTML & "<th><b>" & SDB.Localize("Publisher") & "</b></th>"
	tracklistHTML = tracklistHTML & "<th><b>" & SDB.Localize("Discogs Release") & "</b></th>"
	tracklistHTML = tracklistHTML & "<th><b>" & SDB.Localize("Catalog") & "</b></th>"
	tracklistHTML = tracklistHTML & "<th><b>" & SDB.Localize("Format") & "</b></th>"
	tracklistHTML = tracklistHTML & "<th><b>" & SDB.Localize("Country") & "</b></th>"
	tracklistHTML = tracklistHTML & "<th><b>" & SDB.Localize("Comment") & "</b></th>"
	tracklistHTML = tracklistHTML & "</tr>"
	
	j = 0
	Dim itm2, tmpYear, cSubTrack, subTrackTitle, position
	
	'First Line Table (Original Track tags)
	For c = 0 To NewTrackList.Count - 1
		
		Set itm2 = NewTrackList.Item(c)
		tracklistHTML = tracklistHTML & "<tr bgcolor=""#CCCCCC"">"
		tracklistHTML = tracklistHTML & "<td nowrap align=right><input type=""radio"" id=""" & c & """ name=""trackauswahl""></td><td align=left><b>" & c+1 & "</b></td>"
		
		tracklistHTML = tracklistHTML & "<td align=middle nowrap>" & itm2.TrackOrderStr & "</td>"
		tracklistHTML = tracklistHTML & "<td align=middle nowrap>" & itm2.DiscNumberStr & "</td>"
		tracklistHTML = tracklistHTML & "<td nowrap>" & itm2.Title & "</td>"
		tracklistHTML = tracklistHTML & "<td nowrap>" & itm2.ArtistName & "</td>"
		tracklistHTML = tracklistHTML & "<td nowrap>" & itm2.AlbumName & "</td>"
		tracklistHTML = tracklistHTML & "<td nowrap>" & itm2.AlbumArtistName & "</td>"
		tracklistHTML = tracklistHTML & "<td nowrap>" & itm2.Genre & "</td>"
		
		If itm2.Year  = 0 Then
			tmpYear = ""
		Else
			tmpYear = itm2.Year
		End If
		tracklistHTML = tracklistHTML & "<td nowrap>" & tmpYear & "</td>"
		If itm2.OriginalYear  = -1 Then
			tmpYear = ""
		Else
			tmpYear = itm2.OriginalYear
		End If
		
		tracklistHTML = tracklistHTML & "<td nowrap>" & tmpYear & "</td>"
		tracklistHTML = tracklistHTML & "<td nowrap>" & itm2.Path & "</td>"
		tracklistHTML = tracklistHTML & "<td nowrap>" & itm2.Author & "</td>"
		tracklistHTML = tracklistHTML & "<td nowrap>" & itm2.Conductor & "</td>"
		tracklistHTML = tracklistHTML & "<td nowrap>" & itm2.Producer & "</td>"
		tracklistHTML = tracklistHTML & "<td nowrap>" & itm2.Lyricist & "</td>"
		tracklistHTML = tracklistHTML & "<td nowrap>" & itm2.InvolvedPeople & "</td>"
		tracklistHTML = tracklistHTML & "<td nowrap>" & itm2.Publisher & "</td>"
		
		If ReleaseTag = "Custom1" Then tracklistHTML = tracklistHTML & "<td nowrap>" & itm2.Custom1 & "</td>"
		If ReleaseTag = "Custom2" Then tracklistHTML = tracklistHTML & "<td nowrap>" & itm2.Custom2 & "</td>"
		If ReleaseTag = "Custom3" Then tracklistHTML = tracklistHTML & "<td nowrap>" & itm2.Custom3 & "</td>"
		If ReleaseTag = "Custom4" Then tracklistHTML = tracklistHTML & "<td nowrap>" & itm2.Custom4 & "</td>"
		If ReleaseTag = "Custom5" Then tracklistHTML = tracklistHTML & "<td nowrap>" & itm2.Custom5 & "</td>"
		If ReleaseTag = "Grouping" Then tracklistHTML = tracklistHTML & "<td nowrap>" & itm2.Grouping & "</td>"
		If ReleaseTag = "ISRC" Then tracklistHTML = tracklistHTML & "<td nowrap>" & itm2.ISRC & "</td>"
		If ReleaseTag = "Encoding" Then tracklistHTML = tracklistHTML & "<td nowrap>" & itm2.Encoding & "</td>"
		If ReleaseTag = "Copyright" Then tracklistHTML = tracklistHTML & "<td nowrap>" & itm2.Copyright & "</td>"
		
		If CatalogTag = "Custom1" Then tracklistHTML = tracklistHTML & "<td nowrap>" & itm2.Custom1 & "</td>"
		If CatalogTag = "Custom2" Then tracklistHTML = tracklistHTML & "<td nowrap>" & itm2.Custom2 & "</td>"
		If CatalogTag = "Custom3" Then tracklistHTML = tracklistHTML & "<td nowrap>" & itm2.Custom3 & "</td>"
		If CatalogTag = "Custom4" Then tracklistHTML = tracklistHTML & "<td nowrap>" & itm2.Custom4 & "</td>"
		If CatalogTag = "Custom5" Then tracklistHTML = tracklistHTML & "<td nowrap>" & itm2.Custom5 & "</td>"
		
		If FormatTag = "Custom1" Then tracklistHTML = tracklistHTML & "<td nowrap>" & itm2.Custom1 & "</td>"
		If FormatTag = "Custom2" Then tracklistHTML = tracklistHTML & "<td nowrap>" & itm2.Custom2 & "</td>"
		If FormatTag = "Custom3" Then tracklistHTML = tracklistHTML & "<td nowrap>" & itm2.Custom3 & "</td>"
		If FormatTag = "Custom4" Then tracklistHTML = tracklistHTML & "<td nowrap>" & itm2.Custom4 & "</td>"
		If FormatTag = "Custom5" Then tracklistHTML = tracklistHTML & "<td nowrap>" & itm2.Custom5 & "</td>"
		
		If CountryTag = "Custom1" Then tracklistHTML = tracklistHTML & "<td nowrap>" & itm2.Custom1 & "</td>"
		If CountryTag = "Custom2" Then tracklistHTML = tracklistHTML & "<td nowrap>" & itm2.Custom2 & "</td>"
		If CountryTag = "Custom3" Then tracklistHTML = tracklistHTML & "<td nowrap>" & itm2.Custom3 & "</td>"
		If CountryTag = "Custom4" Then tracklistHTML = tracklistHTML & "<td nowrap>" & itm2.Custom4 & "</td>"
		If CountryTag = "Custom5" Then tracklistHTML = tracklistHTML & "<td nowrap>" & itm2.Custom5 & "</td>"
		
		tracklistHTML = tracklistHTML & "<td nowrap>" & itm2.Comment & "</td>"
		tracklistHTML = tracklistHTML & "</tr>"
		
		cSubTrack = -1
		subTrackTitle = ""
		
		'Subtrack --------------------------------------------------------------------------------------------------------
		If cSubTrack <> -1 And InStr(LCase(position), ".") = 0 Then
			If SubTrackNameSelection = False Then
				Tracks.Item(cSubTrack) = Tracks.Item(cSubTrack) & " (" & subTrackTitle & ")"
			Else
				Tracks.Item(cSubTrack) = subTrackTitle
			End If
			cSubTrack = -1
			subTrackTitle = ""
		End If
		'Subtrack --------------------------------------------------------------------------------------------------------
		
		Do
			If j < tracks.count Then
				If UnselectedTracks(j) = "" Then
					Rem WriteLog "j=" & j
					Rem WriteLog "Tracks.Item(j)=" & Tracks.Item(j)
					
					tracklistHTML = tracklistHTML & "<tr bgcolor=""#FFFFFF"">"
					tracklistHTML = tracklistHTML & "<td></td><td></td>"
					
					If (CheckTrackNum) And TracksNum.Item(j) <> itm2.TrackOrderStr Then
						tracklistHTML = tracklistHTML & "<td bgcolor=""#FFFF00"" align=middle nowrap>" & TracksNum.Item(j) & "</td>"
					Else
						tracklistHTML = tracklistHTML & "<td bgcolor=""#FFFFFF"" align=middle nowrap>" & itm2.TrackOrderStr & "</td>"
					End If
					
					If (CheckDiscNum) And TracksCD.Item(j) <> itm2.DiscNumberStr Then
						tracklistHTML = tracklistHTML & "<td bgcolor=""#FFFF00"" align=middle nowrap>" & TracksCD.Item(j) & "</td>"
					Else
						tracklistHTML = tracklistHTML & "<td bgcolor=""#FFFFFF"" align=middle nowrap>" & itm2.DiscNumberStr & "</td>"
					End If
					
					If StrComp(Tracks.Item(j), itm2.Title, 1) = 0 Then
						tracklistHTML = tracklistHTML & "<td bgcolor=""#FFFFFF"" nowrap>" & itm2.Title & "</td>"
					Else
						tracklistHTML = tracklistHTML & "<td bgcolor=""#FFFF00"" nowrap>" & Tracks.Item(j) & "</td>"
					End If
					
					If (CheckArtist) And ArtistTitles.Item(j) <> itm2.ArtistName Then
						tracklistHTML = tracklistHTML & "<td bgcolor=""#FFFF00"" nowrap>" & ArtistTitles.Item(j) & "</td>"
					Else
						tracklistHTML = tracklistHTML & "<td bgcolor=""#FFFFFF"" nowrap>" & itm2.ArtistName & "</td>"
					End If
					
					REM If (CheckAlbum) And StrComp(AlbumTitle, itm2.AlbumName, 1) = 0 Then
					If (CheckAlbum) And AlbumTitle <> itm2.AlbumName Then
						tracklistHTML = tracklistHTML & "<td bgcolor=""#FFFF00"" nowrap>" & AlbumTitle & "</td>"
					Else
						tracklistHTML = tracklistHTML & "<td bgcolor=""#FFFFFF"" nowrap>" & itm2.AlbumName & "</td>"
					End If
					
					If (CheckAlbumArtist) And AlbumArtist <> itm2.AlbumArtistName Then
						tracklistHTML = tracklistHTML & "<td bgcolor=""#FFFF00"" nowrap>" & AlbumArtist & "</td>"
					Else
						tracklistHTML = tracklistHTML & "<td bgcolor=""#FFFFFF"" nowrap>" & itm2.AlbumArtistName & "</td>"
					End If
					
					If CheckGenre And CheckStyle Then
						GenStyle = Genres & "; " & Styles
						If Genres = "" Then GenStyle = Styles
						If Styles = "" Then GenStyle = Genres
					ElseIf CheckGenre Then
						GenStyle = Genres
					ElseIf CheckStyle Then
						GenStyle = Styles
					End If
					If (CheckGenre Or CheckStyle) And GenStyle <> itm2.Genre Then
						tracklistHTML = tracklistHTML & "<td bgcolor=""#FFFF00"" nowrap>" & GenStyle & "</td>"
					Else
						tracklistHTML = tracklistHTML & "<td bgcolor=""#FFFFFF"" nowrap>" & itm2.Genre & "</td>"
					End If
					
					If itm2.Year = 0 Then
						tmpYear = ""
					Else
						tmpYear = itm2.Year
					End If
					If (CheckDate) And CStr(ReleaseDate) <> CStr(tmpYear) Then
						tracklistHTML = tracklistHTML & "<td bgcolor=""#FFFF00"" nowrap>" & ReleaseDate & "</td>"
					Else
						tracklistHTML = tracklistHTML & "<td bgcolor=""#FFFFFF"" nowrap>" & tmpYear & "</td>"
					End If
					
					If itm2.OriginalYear = -1 Then
						tmpYear = ""
					Else
						tmpYear = itm2.OriginalYear
					End If
					If (CheckOrigDate) And CStr(OriginalDate) <> CStr(tmpYear) Then
						tracklistHTML = tracklistHTML & "<td bgcolor=""#FFFF00"" nowrap>" & OriginalDate & "</td>"
					Else
						tracklistHTML = tracklistHTML & "<td bgcolor=""#FFFFFF"" nowrap>" & tmpYear & "</td>"
					End If
					
					tracklistHTML = tracklistHTML & "<td bgcolor=""#FFFFFF"" nowrap>" & itm2.Path & "</td>"
					
					If (CheckComposer) And Composers.Item(j) <> itm2.Author Then
						tracklistHTML = tracklistHTML & "<td bgcolor=""#FFFF00"" nowrap>" & Composers.Item(j) & "</td>"
					Else
						tracklistHTML = tracklistHTML & "<td bgcolor=""#FFFFFF"" nowrap>" & itm2.Author & "</td>"
					End If
					
					If (CheckConductor) And Conductors.Item(j) <> itm2.Conductor Then
						tracklistHTML = tracklistHTML & "<td bgcolor=""#FFFF00"" nowrap>" & Conductors.Item(j) & "</td>"
					Else
						tracklistHTML = tracklistHTML & "<td bgcolor=""#FFFFFF"" nowrap>" & itm2.Conductor & "</td>"
					End If
					
					If (CheckProducer) And Producers.Item(j) <> itm2.Producer Then
						tracklistHTML = tracklistHTML & "<td bgcolor=""#FFFF00"" nowrap>" & Producers.Item(j) & "</td>"
					Else
						tracklistHTML = tracklistHTML & "<td bgcolor=""#FFFFFF"" nowrap>" & itm2.Producer & "</td>"
					End If
					
					If (CheckLyricist) And Lyricists.Item(j) <> itm2.Lyricist Then
						tracklistHTML = tracklistHTML & "<td bgcolor=""#FFFF00"" nowrap>" & Lyricists.Item(j) & "</td>"
					Else
						tracklistHTML = tracklistHTML & "<td bgcolor=""#FFFFFF"" nowrap>" & itm2.Lyricist & "</td>"
					End If
					
					If (CheckInvolved) And InvolvedArtists.Item(j) <> itm2.InvolvedPeople Then
						tracklistHTML = tracklistHTML & "<td bgcolor=""#FFFF00"" nowrap>" & InvolvedArtists.Item(j) & "</td>"
					Else
						tracklistHTML = tracklistHTML & "<td bgcolor=""#FFFFFF"" nowrap>" & itm2.InvolvedPeople & "</td>"
					End If
					
					If (CheckLabel) And theLabels <> itm2.Publisher Then
						tracklistHTML = tracklistHTML & "<td bgcolor=""#FFFF00"" nowrap>" & theLabels & "</td>"
					Else
						tracklistHTML = tracklistHTML & "<td bgcolor=""#FFFFFF"" nowrap>" & itm2.Publisher & "</td>"
					End If
					
					If ReleaseTag = "Custom1" Then
						If itm2.Custom1 <> CurrentReleaseID Then
							tracklistHTML = tracklistHTML & "<td bgcolor=""#FFFF00"" nowrap>" & CurrentReleaseID & "</td>"
						Else
							tracklistHTML = tracklistHTML & "<td bgcolor=""#FFFFFF"" nowrap>" & itm2.Custom1 & "</td>"
						End If
					End If
					If ReleaseTag = "Custom2" Then
						If itm2.Custom2 <> CurrentReleaseID Then
							tracklistHTML = tracklistHTML & "<td bgcolor=""#FFFF00"" nowrap>" & CurrentReleaseID & "</td>"
						Else
							tracklistHTML = tracklistHTML & "<td bgcolor=""#FFFFFF"" nowrap>" & itm2.Custom2 & "</td>"
						End If
					End If
					If ReleaseTag = "Custom3" Then
						If itm2.Custom3 <> CurrentReleaseID Then
							tracklistHTML = tracklistHTML & "<td bgcolor=""#FFFF00"" nowrap>" & CurrentReleaseID & "</td>"
						Else
							tracklistHTML = tracklistHTML & "<td bgcolor=""#FFFFFF"" nowrap>" & itm2.Custom3 & "</td>"
						End If
					End If
					If ReleaseTag = "Custom4" Then
						If itm2.Custom4 <> CurrentReleaseID Then
							tracklistHTML = tracklistHTML & "<td bgcolor=""#FFFF00"" nowrap>" & CurrentReleaseID & "</td>"
						Else
							tracklistHTML = tracklistHTML & "<td bgcolor=""#FFFFFF"" nowrap>" & itm2.Custom4 & "</td>"
						End If
					End If
					If ReleaseTag = "Custom5" Then
						If itm2.Custom5 <> CurrentReleaseID Then
							tracklistHTML = tracklistHTML & "<td bgcolor=""#FFFF00"" nowrap>" & CurrentReleaseID & "</td>"
						Else
							tracklistHTML = tracklistHTML & "<td bgcolor=""#FFFFFF"" nowrap>" & itm2.Custom5 & "</td>"
						End If
					End If
					
					If CatalogTag = "Custom1" Then
						If itm2.Custom1 <> theCatalogs Then
							tracklistHTML = tracklistHTML & "<td bgcolor=""#FFFF00"" nowrap>" & theCatalogs & "</td>"
						Else
							tracklistHTML = tracklistHTML & "<td bgcolor=""#FFFFFF"" nowrap>" & itm2.Custom1 & "</td>"
						End If
					End If
					If CatalogTag = "Custom2" Then
						If itm2.Custom2 <> theCatalogs Then
							tracklistHTML = tracklistHTML & "<td bgcolor=""#FFFF00"" nowrap>" & theCatalogs & "</td>"
						Else
							tracklistHTML = tracklistHTML & "<td bgcolor=""#FFFFFF"" nowrap>" & itm2.Custom2 & "</td>"
						End If
					End If
					If CatalogTag = "Custom3" Then
						If itm2.Custom3 <> theCatalogs Then
							tracklistHTML = tracklistHTML & "<td bgcolor=""#FFFF00"" nowrap>" & theCatalogs & "</td>"
						Else
							tracklistHTML = tracklistHTML & "<td bgcolor=""#FFFFFF"" nowrap>" & itm2.Custom3 & "</td>"
						End If
					End If
					If CatalogTag = "Custom4" Then
						If itm2.Custom4 <> theCatalogs Then
							tracklistHTML = tracklistHTML & "<td bgcolor=""#FFFF00"" nowrap>" & theCatalogs & "</td>"
						Else
							tracklistHTML = tracklistHTML & "<td bgcolor=""#FFFFFF"" nowrap>" & itm2.Custom4 & "</td>"
						End If
					End If
					If CatalogTag = "Custom5" Then
						If itm2.Custom5 <> theCatalogs Then
							tracklistHTML = tracklistHTML & "<td bgcolor=""#FFFF00"" nowrap>" & theCatalogs & "</td>"
						Else
							tracklistHTML = tracklistHTML & "<td bgcolor=""#FFFFFF"" nowrap>" & itm2.Custom5 & "</td>"
						End If
					End If
					
					If FormatTag = "Custom1" Then
						If itm2.Custom1 <> theFormat Then
							tracklistHTML = tracklistHTML & "<td bgcolor=""#FFFF00"" nowrap>" & theFormat & "</td>"
						Else
							tracklistHTML = tracklistHTML & "<td bgcolor=""#FFFFFF"" nowrap>" & itm2.Custom1 & "</td>"
						End If
					End If
					If FormatTag = "Custom2" Then
						If itm2.Custom2 <> theFormat Then
							tracklistHTML = tracklistHTML & "<td bgcolor=""#FFFF00"" nowrap>" & theFormat & "</td>"
						Else
							tracklistHTML = tracklistHTML & "<td bgcolor=""#FFFFFF"" nowrap>" & itm2.Custom2 & "</td>"
						End If
					End If
					If FormatTag = "Custom3" Then
						If itm2.Custom3 <> theFormat Then
							tracklistHTML = tracklistHTML & "<td bgcolor=""#FFFF00"" nowrap>" & theFormat & "</td>"
						Else
							tracklistHTML = tracklistHTML & "<td bgcolor=""#FFFFFF"" nowrap>" & itm2.Custom3 & "</td>"
						End If
					End If
					If FormatTag = "Custom4" Then
						If itm2.Custom4 <> theFormat Then
							tracklistHTML = tracklistHTML & "<td bgcolor=""#FFFF00"" nowrap>" & theFormat & "</td>"
						Else
							tracklistHTML = tracklistHTML & "<td bgcolor=""#FFFFFF"" nowrap>" & itm2.Custom4 & "</td>"
						End If
					End If
					If FormatTag = "Custom5" Then
						If itm2.Custom5 <> theFormat Then
							tracklistHTML = tracklistHTML & "<td bgcolor=""#FFFF00"" nowrap>" & theFormat & "</td>"
						Else
							tracklistHTML = tracklistHTML & "<td bgcolor=""#FFFFFF"" nowrap>" & itm2.Custom5 & "</td>"
						End If
					End If
					
					If CountryTag = "Custom1" Then
						If itm2.Custom1 <> theCountry Then
							tracklistHTML = tracklistHTML & "<td bgcolor=""#FFFF00"" nowrap>" & theCountry & "</td>"
						Else
							tracklistHTML = tracklistHTML & "<td bgcolor=""#FFFFFF"" nowrap>" & itm2.Custom1 & "</td>"
						End If
					End If
					If CountryTag = "Custom2" Then
						If itm2.Custom2 <> theCountry Then
							tracklistHTML = tracklistHTML & "<td bgcolor=""#FFFF00"" nowrap>" & theCountry & "</td>"
						Else
							tracklistHTML = tracklistHTML & "<td bgcolor=""#FFFFFF"" nowrap>" & itm2.Custom2 & "</td>"
						End If
					End If
					If CountryTag = "Custom3" Then
						If itm2.Custom3 <> theCountry Then
							tracklistHTML = tracklistHTML & "<td bgcolor=""#FFFF00"" nowrap>" & theCountry & "</td>"
						Else
							tracklistHTML = tracklistHTML & "<td bgcolor=""#FFFFFF"" nowrap>" & itm2.Custom3 & "</td>"
						End If
					End If
					If CountryTag = "Custom4" Then
						If itm2.Custom4 <> theCountry Then
							tracklistHTML = tracklistHTML & "<td bgcolor=""#FFFF00"" nowrap>" & theCountry & "</td>"
						Else
							tracklistHTML = tracklistHTML & "<td bgcolor=""#FFFFFF"" nowrap>" & itm2.Custom4 & "</td>"
						End If
					End If
					If CountryTag = "Custom5" Then
						If itm2.Custom5 <> theCountry Then
							tracklistHTML = tracklistHTML & "<td bgcolor=""#FFFF00"" nowrap>" & theCountry & "</td>"
						Else
							tracklistHTML = tracklistHTML & "<td bgcolor=""#FFFFFF"" nowrap>" & itm2.Custom5 & "</td>"
						End If
					End If
					
					
					If (CheckComment) And Comment <> itm2.Comment Then
						tracklistHTML = tracklistHTML & "<td bgcolor=""#FFFF00"" nowrap>" & Comment & "</td>"
						Rem WriteLog "Comment_DB=" & itm2.Comment
						Rem WriteLog "Comment_Discogs=" & Comment
					Else
						tracklistHTML = tracklistHTML & "<td bgcolor=""#FFFFFF"" nowrap>" & itm2.Comment & "</td>"
					End If
					
					tracklistHTML = tracklistHTML & "</tr>"
					j = j + 1
					Exit Do
				Else
					j = j + 1
				End If
			Else
				tracklistHTML = tracklistHTML & "<tr bgcolor=""#CCCCCC"">"
				tracklistHTML = tracklistHTML & "<td></td><td>?</td>"
				tracklistHTML = tracklistHTML & "<td align=middle nowrap>" & itm2.TrackOrderStr & "</td>"
				tracklistHTML = tracklistHTML & "<td align=middle nowrap>" & itm2.DiscNumberStr & "</td>"
				tracklistHTML = tracklistHTML & "<td nowrap>" & itm2.Title & "</td>"
				tracklistHTML = tracklistHTML & "<td nowrap>" & itm2.ArtistName & "</td>"
				tracklistHTML = tracklistHTML & "<td nowrap>" & itm2.AlbumName & "</td>"
				tracklistHTML = tracklistHTML & "<td nowrap>" & itm2.AlbumArtistName & "</td>"
				tracklistHTML = tracklistHTML & "<td nowrap>" & itm2.Genre & "</td>"
				tracklistHTML = tracklistHTML & "<td nowrap>" & itm2.Year & "</td>"
				tracklistHTML = tracklistHTML & "<td nowrap>" & itm2.OriginalYear & "</td>"
				tracklistHTML = tracklistHTML & "<td nowrap>" & itm2.Path & "</td>"
				tracklistHTML = tracklistHTML & "<td nowrap>" & itm2.Author & "</td>"
				tracklistHTML = tracklistHTML & "<td nowrap>" & itm2.Conductor & "</td>"
				tracklistHTML = tracklistHTML & "<td nowrap>" & itm2.Producer & "</td>"
				tracklistHTML = tracklistHTML & "<td nowrap>" & itm2.Lyricist & "</td>"
				tracklistHTML = tracklistHTML & "<td nowrap>" & itm2.InvolvedPeople & "</td>"
				tracklistHTML = tracklistHTML & "<td nowrap>" & itm2.Publisher & "</td>"
				
				If ReleaseTag = "Custom1" Then tracklistHTML = tracklistHTML & "<td nowrap>" & itm2.Custom1 & "</td>"
				If ReleaseTag = "Custom2" Then tracklistHTML = tracklistHTML & "<td nowrap>" & itm2.Custom2 & "</td>"
				If ReleaseTag = "Custom3" Then tracklistHTML = tracklistHTML & "<td nowrap>" & itm2.Custom3 & "</td>"
				If ReleaseTag = "Custom4" Then tracklistHTML = tracklistHTML & "<td nowrap>" & itm2.Custom4 & "</td>"
				If ReleaseTag = "Custom5" Then tracklistHTML = tracklistHTML & "<td nowrap>" & itm2.Custom5 & "</td>"
				
				If CatalogTag = "Custom1" Then tracklistHTML = tracklistHTML & "<td nowrap>" & itm2.Custom1 & "</td>"
				If CatalogTag = "Custom2" Then tracklistHTML = tracklistHTML & "<td nowrap>" & itm2.Custom2 & "</td>"
				If CatalogTag = "Custom3" Then tracklistHTML = tracklistHTML & "<td nowrap>" & itm2.Custom3 & "</td>"
				If CatalogTag = "Custom4" Then tracklistHTML = tracklistHTML & "<td nowrap>" & itm2.Custom4 & "</td>"
				If CatalogTag = "Custom5" Then tracklistHTML = tracklistHTML & "<td nowrap>" & itm2.Custom5 & "</td>"
				
				If FormatTag = "Custom1" Then tracklistHTML = tracklistHTML & "<td nowrap>" & itm2.Custom1 & "</td>"
				If FormatTag = "Custom2" Then tracklistHTML = tracklistHTML & "<td nowrap>" & itm2.Custom2 & "</td>"
				If FormatTag = "Custom3" Then tracklistHTML = tracklistHTML & "<td nowrap>" & itm2.Custom3 & "</td>"
				If FormatTag = "Custom4" Then tracklistHTML = tracklistHTML & "<td nowrap>" & itm2.Custom4 & "</td>"
				If FormatTag = "Custom5" Then tracklistHTML = tracklistHTML & "<td nowrap>" & itm2.Custom5 & "</td>"
				
				If CountryTag = "Custom1" Then tracklistHTML = tracklistHTML & "<td nowrap>" & itm2.Custom1 & "</td>"
				If CountryTag = "Custom2" Then tracklistHTML = tracklistHTML & "<td nowrap>" & itm2.Custom2 & "</td>"
				If CountryTag = "Custom3" Then tracklistHTML = tracklistHTML & "<td nowrap>" & itm2.Custom3 & "</td>"
				If CountryTag = "Custom4" Then tracklistHTML = tracklistHTML & "<td nowrap>" & itm2.Custom4 & "</td>"
				If CountryTag = "Custom5" Then tracklistHTML = tracklistHTML & "<td nowrap>" & itm2.Custom5 & "</td>"
				
				tracklistHTML = tracklistHTML & "<td nowrap>" & itm2.Comment & "</td>"
				tracklistHTML = tracklistHTML & "</tr>"
				Exit Do
			End If
		Loop While True
	Next
	
	tracklistHTML = tracklistHTML & "</table>"
	tracklistHTML = tracklistHTML & "</body>"
	tracklistHTML = tracklistHTML & "</HTML>"
	
	Dim Btn2 : Set Btn2 = SDB.Objects("Btn2")
	If InStr(tracklistHTML, "#FFFF00") = 0 Then
		Rem Btn1.Common.Visible = true
		Btn2.Common.Enabled = False
		Btn2.Common.Hint = "There are no different content in the tags"
		trackliste_aufbauen = True
	Else
		Rem Btn1.Common.Visible = false
		Btn2.Common.Enabled = True
		Btn2.Common.Hint = "Update the tag(s) with different content"
		trackliste_aufbauen = False
	End If
	WriteLog "Stop Trackliste_aufbauen"
	
End Function


Sub ShowCountryFilter
	
	Dim Form, iWidth, CountColumn, filterHTML, filterHTMLDoc, WebBrowser2, countrybutton, FilterFound
	Dim i, a
	Set Form = UI.NewForm
	Form.Common.Width = 675
	Form.Common.Height = 600
	Form.FormPosition = 4
	Form.Caption = "Choose the country's to search for"
	Form.BorderStyle = 3
	Form.StayOnTop = True
	SDB.Objects("FilterForm") = Form
	SDB.Objects("Filter") = CountryList
	CountColumn = (CountryList.Count - 1) / 94
	iWidth = (CountryList.Count - 1) / 94 * 200
	filterHTML = GetFilterHTML(iWidth, 93, CountColumn)
	
	Dim Foot : Set Foot = SDB.UI.NewPanel(Form)
	Foot.Common.Align = 2
	Foot.Common.Height = 35
	
	Dim Btn : Set Btn = SDB.UI.NewButton(Foot)
	Btn.Caption = SDB.Localize("Cancel")
	Btn.Common.Width = 85
	Btn.Common.Height = 25
	Btn.Common.Left = Form.Common.Width - Btn.Common.Width - 30
	Btn.Common.Top = 6
	Btn.Common.Anchors = 2+4
	Btn.UseScript = Script.ScriptPath
	Btn.ModalResult = 2
	Btn.Cancel = True
	
	Dim Btn2 : Set Btn2 = SDB.UI.NewButton(Foot)
	Btn2.Caption = SDB.Localize("Ok")
	Btn2.Common.Width = 85
	Btn2.Common.Height = 25
	Btn2.Common.Left = Btn.Common.Left - Btn2.Common.Width - 5
	Btn2.Common.Top = 6
	Btn2.Common.Anchors = 2+4
	Btn2.UseScript = Script.ScriptPath
	Btn2.ModalResult = 1
	Btn2.Default = True
	
	Dim Btn3 : Set Btn3 = SDB.UI.NewButton(Foot)
	Btn3.Caption = SDB.Localize("&Check All")
	Btn3.Common.Width = 85
	Btn3.Common.Height = 25
	Btn3.Common.Left = 15
	Btn3.Common.Top = 6
	Btn3.Common.Anchors = 2+4
	Script.RegisterEvent Btn3, "OnClick", "Btn6Click"
	
	Dim Btn4 : Set Btn4 = SDB.UI.NewButton(Foot)
	Btn4.Caption = SDB.Localize("&Uncheck all")
	Btn4.Common.Width = 85
	Btn4.Common.Height = 25
	Btn4.Common.Left = Btn3.Common.Left + Btn4.Common.Width + 5
	Btn4.Common.Top = 6
	Btn4.Common.Anchors = 2+4
	Script.RegisterEvent Btn4, "OnClick", "Btn7Click"
	
	Set WebBrowser2 = UI.NewActiveX(Form, "Shell.Explorer")
	WebBrowser2.Common.Align = 5
	WebBrowser2.Common.ControlName = "WebBrowser2"
	WebBrowser2.Common.Top = 100
	WebBrowser2.Common.Left = 100
	
	SDB.Objects("WebBrowser2") = WebBrowser2
	WebBrowser2.Interf.Visible = True
	WebBrowser2.Common.BringToFront
	
	WebBrowser2.SetHTMLDocument filterHTML
	Set filterHTMLDoc = WebBrowser2.Interf.Document
	
	For i = 1 To CountryList.Count - 1
		Set countrybutton = filterHTMLDoc.getElementById("Filter" & i)
		If CountryFilterList.Item(i) = "1" Then
			countrybutton.checked = True
		End If
	Next
	
	If Form.ShowModal = 1 Then
		FilterFound = False
		For a = 1 To CountryList.Count - 1
			Set countrybutton = filterHTMLDoc.getElementById("Filter" & a)
			If countrybutton.checked = True Then
				CountryFilterList.Item(a) = "1"
				FilterFound = True
			Else
				CountryFilterList.Item(a) = "0"
			End If
		Next
		If FilterFound = False Then
			FilterCountry = "None"
			CountryFilterList.Item(0) = "0"
		Else
			FilterCountry = "Use Country Filter"
			CountryFilterList.Item(0) = "1"
		End If
		SDB.Objects("WebBrowser2") = Nothing
		SDB.Objects("FilterForm") = Nothing
		SDB.Objects("Filter") = Nothing
		FindResults SavedSearchTerm, "", ""
	Else
		SDB.Objects("WebBrowser2") = Nothing
		SDB.Objects("FilterForm") = Nothing
		SDB.Objects("Filter") = Nothing
	End If
	
End Sub


Sub ShowMediaFormatFilter
	
	Dim Form, iWidth, CountColumn, filterHTML, filterHTMLDoc, WebBrowser2, MediaFormatButton, FilterFound
	Dim i, a
	Set Form = UI.NewForm
	Form.Common.Width = 380
	Form.Common.Height = 700
	Form.FormPosition = 4
	Form.Caption = "Choose the MediaFormat to search for"
	Form.BorderStyle = 3
	Form.StayOnTop = True
	SDB.Objects("FilterForm") = Form
	SDB.Objects("Filter") = MediaFormatList
	iWidth = (MediaFormatList.Count - 1) / 24 * 150
	CountColumn = (MediaFormatList.Count - 1) / 24
	
	filterHTML = GetFilterHTML(iWidth, 23, CountColumn)
	
	Dim Foot : Set Foot = SDB.UI.NewPanel(Form)
	Foot.Common.Align = 2
	Foot.Common.Height = 35
	
	Dim Btn : Set Btn = SDB.UI.NewButton(Foot)
	Btn.Caption = SDB.Localize("Cancel")
	Btn.Common.Width = 85
	Btn.Common.Height = 25
	Btn.Common.Left = Form.Common.Width - Btn.Common.Width - 20
	Btn.Common.Top = 6
	Btn.Common.Anchors = 2+4
	Btn.UseScript = Script.ScriptPath
	Btn.ModalResult = 2
	Btn.Cancel = True
	
	Dim Btn2 : Set Btn2 = SDB.UI.NewButton(Foot)
	Btn2.Caption = SDB.Localize("Ok")
	Btn2.Common.Width = 85
	Btn2.Common.Height = 25
	Btn2.Common.Left = Btn.Common.Left - Btn2.Common.Width - 5
	Btn2.Common.Top = 6
	Btn2.Common.Anchors = 2+4
	Btn2.UseScript = Script.ScriptPath
	Btn2.ModalResult = 1
	Btn2.Default = True
	
	Dim Btn3 : Set Btn3 = SDB.UI.NewButton(Foot)
	Btn3.Caption = SDB.Localize("&Check All")
	Btn3.Common.Width = 85
	Btn3.Common.Height = 25
	Btn3.Common.Left = 5
	Btn3.Common.Top = 6
	Btn3.Common.Anchors = 2+4
	Script.RegisterEvent Btn3, "OnClick", "Btn6Click"
	
	Dim Btn4 : Set Btn4 = SDB.UI.NewButton(Foot)
	Btn4.Caption = SDB.Localize("&Uncheck all")
	Btn4.Common.Width = 85
	Btn4.Common.Height = 25
	Btn4.Common.Left = Btn3.Common.Left + Btn4.Common.Width + 5
	Btn4.Common.Top = 6
	Btn4.Common.Anchors = 2+4
	Script.RegisterEvent Btn4, "OnClick", "Btn7Click"
	
	Set WebBrowser2 = UI.NewActiveX(Form, "Shell.Explorer")
	WebBrowser2.Common.Align = 5
	WebBrowser2.Common.ControlName = "WebBrowser2"
	WebBrowser2.Common.Top = 100
	WebBrowser2.Common.Left = 100
	
	SDB.Objects("WebBrowser2") = WebBrowser2
	WebBrowser2.Interf.Visible = True
	WebBrowser2.Common.BringToFront
	
	WebBrowser2.SetHTMLDocument filterHTML
	Set filterHTMLDoc = WebBrowser2.Interf.Document
	
	For i = 1 To MediaFormatList.Count - 1
		Set MediaFormatButton = filterHTMLDoc.getElementById("Filter" & i)
		If MediaFormatFilterList.Item(i) = "1" Then
			MediaFormatButton.checked = True
		End If
	Next
	
	If Form.ShowModal = 1 Then
		FilterFound = False
		For a = 1 To MediaFormatList.Count - 1
			Set MediaFormatButton = filterHTMLDoc.getElementById("Filter" & a)
			If MediaFormatButton.checked = True Then
				MediaFormatFilterList.Item(a) = "1"
				FilterFound = True
			Else
				MediaFormatFilterList.Item(a) = "0"
			End If
		Next
		If FilterFound = False Then
			FilterMediaFormat = "None"
			MediaFormatFilterList.Item(0) = "0"
		Else
			FilterMediaFormat = "Use MediaFormat Filter"
			MediaFormatFilterList.Item(0) = "1"
		End If
		SDB.Objects("WebBrowser2") = Nothing
		SDB.Objects("FilterForm") = Nothing
		SDB.Objects("Filter") = Nothing
		FindResults SavedSearchTerm, "", ""
	Else
		SDB.Objects("WebBrowser2") = Nothing
		SDB.Objects("FilterForm") = Nothing
		SDB.Objects("Filter") = Nothing
	End If
	
End Sub


Sub ShowMediaTypeFilter
	
	Dim Form, iWidth, CountColumn, filterHTML, filterHTMLDoc, WebBrowser2, MediaTypeButton, FilterFound
	Dim i, a
	Set Form = UI.NewForm
	Form.Common.Width = 420
	Form.Common.Height = 600
	Form.FormPosition = 4
	Form.Caption = "Choose the MediaType to search for"
	Form.BorderStyle = 3
	Form.StayOnTop = True
	SDB.Objects("FilterForm") = Form
	SDB.Objects("Filter") = MediaTypeList
	iWidth = (MediaTypeList.Count - 1) / 19 * 175
	CountColumn = (MediaTypeList.Count - 1) / 19
	
	filterHTML = GetFilterHTML(iWidth, 18, CountColumn)
	
	Dim Foot : Set Foot = SDB.UI.NewPanel(Form)
	Foot.Common.Align = 2
	Foot.Common.Height = 35
	
	Dim Btn : Set Btn = SDB.UI.NewButton(Foot)
	Btn.Caption = SDB.Localize("Cancel")
	Btn.Common.Width = 85
	Btn.Common.Height = 25
	Btn.Common.Left = Form.Common.Width - Btn.Common.Width - 30
	Btn.Common.Top = 6
	Btn.Common.Anchors = 2+4
	Btn.UseScript = Script.ScriptPath
	Btn.ModalResult = 2
	Btn.Cancel = True
	
	Dim Btn2 : Set Btn2 = SDB.UI.NewButton(Foot)
	Btn2.Caption = SDB.Localize("Ok")
	Btn2.Common.Width = 85
	Btn2.Common.Height = 25
	Btn2.Common.Left = Btn.Common.Left - Btn2.Common.Width - 5
	Btn2.Common.Top = 6
	Btn2.Common.Anchors = 2+4
	Btn2.UseScript = Script.ScriptPath
	Btn2.ModalResult = 1
	Btn2.Default = True
	
	Dim Btn3 : Set Btn3 = SDB.UI.NewButton(Foot)
	Btn3.Caption = SDB.Localize("&Check All")
	Btn3.Common.Width = 85
	Btn3.Common.Height = 25
	Btn3.Common.Left = 15
	Btn3.Common.Top = 6
	Btn3.Common.Anchors = 2+4
	Script.RegisterEvent Btn3, "OnClick", "Btn6Click"
	
	Dim Btn4 : Set Btn4 = SDB.UI.NewButton(Foot)
	Btn4.Caption = SDB.Localize("&Uncheck all")
	Btn4.Common.Width = 85
	Btn4.Common.Height = 25
	Btn4.Common.Left = Btn3.Common.Left + Btn4.Common.Width + 5
	Btn4.Common.Top = 6
	Btn4.Common.Anchors = 2+4
	Script.RegisterEvent Btn4, "OnClick", "Btn7Click"
	
	Set WebBrowser2 = UI.NewActiveX(Form, "Shell.Explorer")
	WebBrowser2.Common.Align = 5
	WebBrowser2.Common.ControlName = "WebBrowser2"
	WebBrowser2.Common.Top = 100
	WebBrowser2.Common.Left = 100
	
	SDB.Objects("WebBrowser2") = WebBrowser2
	WebBrowser2.Interf.Visible = True
	WebBrowser2.Common.BringToFront
	
	WebBrowser2.SetHTMLDocument filterHTML
	Set filterHTMLDoc = WebBrowser2.Interf.Document
	
	For i = 1 To MediaTypeList.Count - 1
		Set MediaTypeButton = filterHTMLDoc.getElementById("Filter" & i)
		If MediaTypeFilterList.Item(i) = "1" Then
			MediaTypeButton.checked = True
		End If
	Next
	
	If Form.ShowModal = 1 Then
		FilterFound = False
		For a = 1 To MediaTypeList.Count - 1
			Set MediaTypeButton = filterHTMLDoc.getElementById("Filter" & a)
			If MediaTypeButton.checked = True Then
				MediaTypeFilterList.Item(a) = "1"
				FilterFound = True
			Else
				MediaTypeFilterList.Item(a) = "0"
			End If
		Next
		If FilterFound = False Then
			FilterMediaType = "None"
			MediaTypeFilterList.Item(0) = "0"
		Else
			FilterMediaType = "Use MediaType Filter"
			MediaTypeFilterList.Item(0) = "1"
		End If
		SDB.Objects("WebBrowser2") = Nothing
		SDB.Objects("FilterForm") = Nothing
		SDB.Objects("Filter") = Nothing
		FindResults SavedSearchTerm, "", ""
	Else
		SDB.Objects("WebBrowser2") = Nothing
		SDB.Objects("FilterForm") = Nothing
		SDB.Objects("Filter") = Nothing
	End If
	
End Sub


Sub ShowYearFilter
	
	Dim Form, iWidth, CountColumn, filterHTML, filterHTMLDoc, YearButton, FilterFound
	Dim i, a, row
	Set Form = UI.NewForm
	Form.Common.Width = 550
	Form.Common.Height = 550
	Form.FormPosition = 4
	Form.Caption = "Choose the Year to search for"
	Form.BorderStyle = 3
	Form.StayOnTop = True
	SDB.Objects("FilterForm") = Form
	SDB.Objects("Filter") = YearList
	'CountColumn = 6
	If ((YearList.Count - 1) / 6) = Int((YearList.Count - 1) / 6) Then
		iWidth = (YearList.Count - 1) / 6 * 25
		row = Int((YearList.Count - 1) / 6)
		CountColumn = 6
	Else
		row = Int((YearList.Count - 1) / 6) + 1
		iWidth = (YearList.Count - 1) / 6 * 25
		CountColumn = 6
	End If
	
	filterHTML = GetFilterHTML(iWidth, row-1, CountColumn)
	
	Dim Foot : Set Foot = SDB.UI.NewPanel(Form)
	Foot.Common.Align = 2
	Foot.Common.Height = 35
	
	Dim Btn : Set Btn = SDB.UI.NewButton(Foot)
	Btn.Caption = SDB.Localize("Cancel")
	Btn.Common.Width = 85
	Btn.Common.Height = 25
	Btn.Common.Left = Form.Common.Width - Btn.Common.Width - 30
	Btn.Common.Top = 6
	Btn.Common.Anchors = 2+4
	Btn.UseScript = Script.ScriptPath
	Btn.ModalResult = 2
	Btn.Cancel = True
	
	Dim Btn2 : Set Btn2 = SDB.UI.NewButton(Foot)
	Btn2.Caption = SDB.Localize("Ok")
	Btn2.Common.Width = 85
	Btn2.Common.Height = 25
	Btn2.Common.Left = Btn.Common.Left - Btn2.Common.Width - 5
	Btn2.Common.Top = 6
	Btn2.Common.Anchors = 2+4
	Btn2.UseScript = Script.ScriptPath
	Btn2.ModalResult = 1
	Btn2.Default = True
	
	Dim Btn3 : Set Btn3 = SDB.UI.NewButton(Foot)
	Btn3.Caption = SDB.Localize("&Check All")
	Btn3.Common.Width = 85
	Btn3.Common.Height = 25
	Btn3.Common.Left = 15
	Btn3.Common.Top = 6
	Btn3.Common.Anchors = 2+4
	Script.RegisterEvent Btn3, "OnClick", "Btn6Click"
	
	Dim Btn4 : Set Btn4 = SDB.UI.NewButton(Foot)
	Btn4.Caption = SDB.Localize("&Uncheck all")
	Btn4.Common.Width = 85
	Btn4.Common.Height = 25
	Btn4.Common.Left = Btn3.Common.Left + Btn4.Common.Width + 5
	Btn4.Common.Top = 6
	Btn4.Common.Anchors = 2+4
	Script.RegisterEvent Btn4, "OnClick", "Btn7Click"
	
	Set WebBrowser2 = UI.NewActiveX(Form, "Shell.Explorer")
	WebBrowser2.Common.Align = 5
	WebBrowser2.Common.ControlName = "WebBrowser2"
	WebBrowser2.Common.Top = 100
	WebBrowser2.Common.Left = 100
	
	SDB.Objects("WebBrowser2") = WebBrowser2
	WebBrowser2.Interf.Visible = True
	WebBrowser2.Common.BringToFront
	
	WebBrowser2.SetHTMLDocument filterHTML
	Set filterHTMLDoc = WebBrowser2.Interf.Document
	
	For i = 1 To YearList.Count - 1
		Set Yearbutton = filterHTMLDoc.getElementById("Filter" & i)
		If YearFilterList.Item(i) = "1" Then
			Yearbutton.checked = True
		End If
	Next
	
	If Form.ShowModal = 1 Then
		FilterFound = False
		For a = 1 To YearList.Count - 1
			Set Yearbutton = filterHTMLDoc.getElementById("Filter" & a)
			If Yearbutton.checked = True Then
				YearFilterList.Item(a) = "1"
				FilterFound = True
			Else
				YearFilterList.Item(a) = "0"
			End If
		Next
		If FilterFound = False Then
			FilterYear = "None"
			YearFilterList.Item(0) = "0"
		Else
			FilterYear = "Use Year Filter"
			YearFilterList.Item(0) = "1"
		End If
		SDB.Objects("WebBrowser2") = Nothing
		SDB.Objects("FilterForm") = Nothing
		SDB.Objects("Filter") = Nothing
		FindResults SavedSearchTerm, "", ""
	Else
		SDB.Objects("WebBrowser2") = Nothing
		SDB.Objects("FilterForm") = Nothing
		SDB.Objects("Filter") = Nothing
	End If
	
End Sub

Sub Btn6Click
	
	Dim FilterList, filterHTMLDoc, a, filterbutton
	Dim templateHTMLDoc
	Set WebBrowser2 = SDB.Objects("WebBrowser2")
	Set FilterList = SDB.Objects("Filter")
	Set filterHTMLDoc = WebBrowser2.Interf.Document
	For a = 1 To FilterList.Count - 1
		Set filterbutton = filterHTMLDoc.getElementById("Filter" & a)
		filterbutton.checked = True
	Next
	
End Sub


Sub Btn7Click
	
	Dim FilterList, filterHTMLDoc, a, filterbutton
	Dim templateHTMLDoc
	Set WebBrowser2 = SDB.Objects("WebBrowser2")
	Set FilterList = SDB.Objects("Filter")
	Set filterHTMLDoc = WebBrowser2.Interf.Document
	For a = 1 To FilterList.Count - 1
		Set filterbutton = filterHTMLDoc.getElementById("Filter" & a)
		filterbutton.checked = False
	Next
	
End Sub


Function GetFilterHTML(Width, Row, CountColumn)
	
	Dim FilterList, filterHTML, i, a
	Set FilterList = SDB.Objects("Filter")
	filterHTML = "<HTML>"
	filterHTML = filterHTML & "<HEAD>"
	filterHTML = filterHTML & "<style type=""text/css"" media=""screen"">"
	filterHTML = filterHTML & ".tabletext { font-family: Arial, Helvetica, sans-serif; font-size: 8pt;}"
	filterHTML = filterHTML & "</style>"
	filterHTML = filterHTML & "</HEAD>"
	filterHTML = filterHTML & "<table border=0 width=" & Width & " cellspacing=0 cellpadding=1 class=tabletext>"
	For i = 0 To Row
		filterHTML = filterHTML &  "<tr>"
		For a = 1 To CountColumn
			If FilterList.Count = a + (i * CountColumn) Then Exit For
			filterHTML = filterHTML &  "<td><input type=checkbox id=""Filter" & a + (i * CountColumn) & """ >" & FilterList.Item(a + (i * CountColumn))
			filterHTML = filterHTML &  "</td>"
		Next
		filterHTML = filterHTML &  "</tr>"
	Next
	filterHTML = filterHTML &  "</table>"
	filterHTML = filterHTML &  "</body>"
	filterHTML = filterHTML &  "</HTML>"
	GetFilterHTML = filterHTML
	
End Function

Sub Install()

	Dim iniFile : iniFile = SDB.ApplicationPath & "Scripts\Scripts.ini"
	Dim f : Set f = SDB.Tools.IniFileByPath(iniFile)
	If Not (f Is Nothing) Then
		f.StringValue("DiscogsAutoTagWeb_Batch", "Filename") = "DiscogsBatchTagger.vbs"
		f.StringValue("DiscogsAutoTagWeb_Batch", "Procname") = "BatchDiscogsSearch"
		f.StringValue("DiscogsAutoTagWeb_Batch", "Order") = "10"
		f.StringValue("DiscogsAutoTagWeb_Batch", "DisplayName") = "Discogs Batch Tagger"
		f.StringValue("DiscogsAutoTagWeb_Batch", "Description") = "Batch Checks track/album information from discogs.com"
		f.StringValue("DiscogsAutoTagWeb_Batch", "Language") = "VBScript"
		f.StringValue("DiscogsAutoTagWeb_Batch", "ScriptType") = "0"
		SDB.RefreshScriptItems
	End If

End Sub
