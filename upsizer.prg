*==============================================================================
* Program:			Upsizer.prg
* Purpose:			Upsize free tables to SQL Server
* Author:			Doug Hennig
* Last revision:	12/27/2022
*==============================================================================

* Create the Upsizer object.

close databases all
if version(2) = 2
	lcFolder = GetApplicationFolder()
	set path to (lcFolder + 'Source\')
endif version(2) = 2
loUpsize = newobject('Upsizer', 'Upsizer.prg')

* Do the processing.

loUpsize.Execute()
return

*==============================================================================
* Gets the folder the application is in.
*==============================================================================

function GetApplicationFolder
	local laStack[1], ;
		lcProgram, ;
		lcFolder
	if _vfp.StartMode = 0
		astackinfo(laStack)
		lcProgram = laStack[program(-1), 4]
		lcFolder  = addbs(justpath(lcProgram))
	else
		lcFolder = addbs(justpath(_vfp.ServerName))
	endif _vfp.StartMode = 0
	return lcFolder
endfunc

*==============================================================================
* Upsizer class.
*==============================================================================

#define ccCRLF chr(13) + chr(10)
define class Upsizer as Custom
	cApplicationFolder  = ''
		&& the folder this code is running from
	cConfigFile         = ''
		&& the name of the configuration settings file (set to Upsizer.config
		&& in the app folder
	cConfigSettings     = ''
		&& the configuration settings
	cLogFile            = ''
		&& the name of the file containing diagnostic messages
	nLastLog            = 0
		&& the last time we logged something
	nStartLog           = 0
		&& when we started logging
	oEngine               = NULL
		&& a reference to an UpsizingWizard UpsizeEngine object
	nHandle               = 0
		&& the connection handle to SQL Server
	tNextInit             = {//::}
		&& the next time we'll display a message when an upsize process is
		&& started
	tNextUpdate           = {//::}
		&& the next time we'll display a message when an upsize process is
		&& updated
	cTask                 = ''
		&& the current upsizing task
	cTitle                = ''
		&& the current upsizing title
	lAddTablesToDBC       = .T.
		&& .T. to add the tables to the DBC; can set to .F. when debugging to
		&& skip
	lUpsizeDBC            = .T.
		&& .T. to upsize the VFP tables to SQL Server; can set to .F. when
		&& debugging to skip
	lRaiseError           = .T.
		&& .T. to raise an error when something fails to .F. to just log it
	cTempFolder           = ''
		&& folder for temp files

* Settings read from the configuration settings file

	cLogFolder          = ''
		&& the folder containing log files
	lShowProgress         = .F.
		&& .T. to display the progress
	lQuiet                = .F.
		&& .T. to display progress messages
	cConvertedDataFolder  = ''
		&& the location where the free tables exist after being added to a DBC
	cDBCFileName          = ''
		&& the name and path for the upsizing DBC (set to Upsizer.dbc in the
		&& folder specified in ConvertedDataFolder)
	lCreateDatabase       = .F.
		&& .T. to create the database, .F. to upsize to the existing database
	cConnectionString     = ''
		&& the connection string for SQL Server
	cUpsizeFolder         = ''
		&& the folder where the Upsizing Wizard places error files
	lCopySourceTables     = .F.
		&& .T. to copy source tables
	cFreeTablesFolder     = ''
		&& if lCopySourceTables is .T., the folder to copy the source tables
		&& from
	cSQLDatabaseName      = ''
		&& the name of the upsized database
	cCustomSQLScripts     = ''
		&& the name of the folder containing custom SQL scripts to execute
	lDebugMode            = .F.
		&& .T. for debug mode
	lUseSQLBulkCopy       = .F.
		&& .T. to use SQL Bulk Copy
	lUpsizeCharToVarchar  = .F.
		&& .T. to upsize all Character fields to Varchar
	lUpsizeDateToDate     = .F.
		&& .T. to upsize all Date fields to Date (as opposed to DateTime)
	lUpsizeFieldsNullable = .F.
		&& .T. to upsize all fields as nullable except PK, non-nullable FK, and
		&& logical
	lDropExistingDatabase = .F.
		&& .T. to drop an existing database
	lClusteredIndexes     = .F.
		&& .T. to use clustered primary keys
	lCreateDefaults       = .F.
		&& .T. to create defaults for every field
	cBlankDateValue       = ''
		&& the value to use for blank dates
	cDefaultDateValue     = ''
		&& the default to use for dates if lCreateDefaults is .T.
	lCreateReport         = .F.
		&& .T. to create the Upsizing Wizard reports

*==============================================================================
* Set up the environment.
*==============================================================================

	function Init
		on escape
		set asserts on
		set century on
		set cpdialog off
		set deleted on
		set exclusive off
		set hours to 24
		set multilocks on
		set notify off
		set reprocess to 1
		set safety off
		set tableprompt off
		set talk off

* Set the name of the log file.

		This.SetLogFile()

* Get the paths and file names.

		This.cApplicationFolder = GetApplicationFolder()
		This.cConfigFile        = This.cApplicationFolder + 'Upsizer.config'
		This.cTempFolder        = addbs(sys(2023))

* Get the configuration settings.

		This.GetConfigSettings()

* Create the log folder if necessary.

		This.CreateFolder(This.cLogFolder)

* Set the path so we can find other code files.

		if version(2) = 2
			set path to (This.cApplicationFolder + 'Source\') additive
		endif version(2) = 2

* Get the folder for the Upsizing Wizard.

		This.cUpsizeFolder = This.cApplicationFolder + 'Upsize\'

* Get the name and path of the DBC to create.

		This.cDBCFileName = This.cConvertedDataFolder + 'Upsizer.dbc'

* Disable the login dialog.

		sqlsetprop(0, 'DispLogin', 3)
	endfunc

*==============================================================================
* Clean up on exit.
*==============================================================================

	function Destroy
		use in select('UpsizeTables')
		use in select('UpsizeRelations')
		if This.nHandle > 0
			try
				sqldisconnect(This.nHandle)
			catch
			endtry
		endif This.nHandle > 0
	endfunc

*==============================================================================
* Upsize the free tables.
*==============================================================================

	function Execute()
		local lcSQLCommand, ;
			llCancelled, ;
			llOK, ;
			llUpsized, ;
			loException as Exception
		try

* Debug if we're supposed to.

			if This.lDebugMode
				set step on 
			endif This.lDebugMode

* Reset the folders we'll write data to.

			This.ResetFolders()
			This.DisplayStatus('Starting process')

* Open a SQL Server connection.

			This.OpenConnection()

* Open the UpsizeTables and UpsizeRelations tables, which contains the free
* tables to upsize and the relationships between them.

			This.OpenMetaDataTables()

* Copy the free tables if we're supposed to. If not, we won't add tables to the
* database (we're doing a test run with existing tables). Also, don't actually
* copy if the source and converted folders are the same.

			do case
				case This.lCopySourceTables and ;
					This.cFreeTablesFolder = This.cConvertedDataFolder
				case This.lCopySourceTables
					This.DisplayStatus('Copying files')
* RoboCopy flags: /R:2 = 2 retries on failed copies, /W:5 = 5 seconds wait time between retries,
* /MT:8 = multi-threaded using 8 cores.
* Note: can't have spaces in folder names since RoboCopy doesn't like quotes around names
					select UpsizeTables
					scan
						erase (This.cFreeTablesFolder + trim(TABLE) + '.bak')
						erase (This.cFreeTablesFolder + trim(TABLE) + '.tbk')
						text to lcCommand noshow textmerge pretext 2
						robocopy <<This.cFreeTablesFolder>> <<This.cConvertedDataFolder>> <<trim(TABLE)>>.* /R:2 /W:5 /MT:8
						endtext
						run &lcCommand
					endscan
				otherwise
					This.lAddTablesToDBC = .F.
			endcase

* Create the conversion DBC and add the free tables to it.

			do case
				case This.lAddTablesToDBC
					This.DisplayStatus('Creating conversion DBC')
					create database (This.cDBCFileName)
					This.AddTablesToDBC()
					This.DisplayStatus('Validating DBC')
					validate database recover noconsole
				case This.lUpsizeDBC
					open database (This.cDBCFileName) exclusive
			endcase

* Upsize the data.

			if This.lUpsizeDBC
				This.DisplayStatus('Upsizing data')
				This.SetupUpsizingWizard()
				This.SetupUpsizingFields()
				llUpsized = This.Upsize()
			else
				llUpsized = .T.
			endif This.lUpsizeDBC
			if llUpsized

* From this point on, we won't terminate on a SQL error but just log it.

				This.lRaiseError = .F.

* Execute any custom SQL scripts.

				This.ExecuteCustomSQLScripts()

* Create foreign keys constraints.

				if used('UpsizeRelations')
					This.CreateFKConstraints()
				endif used('UpsizeRelations')

* Do any post-upsizing data cleanup.

				This.DataCleanup()
			endif llUpsized

* Log that we're done.

			This.DisplayStatus('Done processing')

* Log any error during processing.

		catch to loException
			This.DisplayStatus('*** Error #' + transform(loException.ErrorNo) + ;
				' occurred in line ' + transform(loException.LineNo) + ;
				' of ' + loException.Procedure + ;
				': ' + loException.Message)

* Disconnect from SQL Server.

		finally
			sqldisconnect(0)
			This.nHandle = 0
			close databases
		endtry

* Cleanup

		if not This.lQuiet
			wait clear
		endif not This.lQuiet
		return
	endfunc

*==============================================================================
* Open meta data tables.
*==============================================================================

	function OpenMetaDataTables
		local lcTable, ;
			laTables[1], ;
			lnTables, ;
			lnI
		This.DisplayStatus('Opening meta data tables')
		lcTable = This.cApplicationFolder + 'UpsizeRelations.dbf'
		if file(lcTable)
			use (This.cApplicationFolder + 'UpsizeRelations') order Main ;
				again shared in 0
		endif file(lcTable)
		lcTable = This.cApplicationFolder + 'UpsizeTables.dbf'
		if file(lcTable)
			select 0
			use (This.cApplicationFolder + 'UpsizeTables') order Table ;
				again shared
		else
			create cursor UpsizeTables (Table C(120), KeyCol C(10), ;
				PostUpsize M)
			index on upper(Table) tag Table
			lnTables = adir(laTables, This.cFreeTablesFolder + '*.dbf')
			for lnI = 1 to lnTables
				insert into UpsizeTables (Table) ;
					values (lower(juststem(laTables[lnI, 1])))
			next lnI
		endif file(lcTable)
	endfunc

*==============================================================================
* Sets the name of the log file.
*==============================================================================

	function SetLogFile
		This.cLogFile = 'Log.txt'
	endfunc

*==============================================================================
* Gets the configuration settings.
*==============================================================================

	function GetConfigSettings()
		local llTimestamp, ;
			llDropExistingDatabase, ;
			lcBlankDateValue
		This.lQuiet                = lower(This.GetConfigSetting('QuietMode')) = 'true'
		This.lShowProgress         = lower(This.GetConfigSetting('ShowProgress')) = 'true'
		This.cLogFolder            = This.GetConfigSetting('LogFolder', .T.)
		llTimestamp                = lower(This.GetConfigSetting('TimestampLogFile')) = 'true'
		This.cLogFile              = This.cLogFolder + juststem(This.cLogFile) + ;
			iif(llTimestamp, ttoc(datetime(), 1), '') + '.txt'
		This.cConnectionString     = This.GetConfigSetting('ConnectionString')
		This.cSQLDatabaseName      = This.GetConfigSetting('SQLDatabaseName')
		This.lCreateDatabase       = lower(This.GetConfigSetting('CreateDatabase')) = 'true'
		llDropExistingDatabase     = lower(This.GetConfigSetting('DropExistingDatabase')) = 'true'
		This.lDropExistingDatabase = This.lCreateDatabase and llDropExistingDatabase
		This.cConvertedDataFolder  = This.GetConfigSetting('ConvertedDataFolder', ;
			.T.)
		This.cFreeTablesFolder     = This.GetConfigSetting('FreeTablesFolder', ;
			.T.)
		This.lCopySourceTables     = lower(This.GetConfigSetting('CopySourceTables')) = 'true'
		This.cCustomSQLScripts     = This.GetConfigSetting('CustomSQLScriptsFolder', ;
			.T.)
		This.lDebugMode            = lower(This.GetConfigSetting('DebugMode')) = 'true'
		This.lCreateReport         = lower(This.GetConfigSetting('CreateReport')) = 'true'
		This.lUseSQLBulkCopy       = lower(This.GetConfigSetting('SQLBulkCopy')) = 'true'
		This.lUpsizeCharToVarchar  = lower(This.GetConfigSetting('UpsizeCharToVarchar')) = 'true'
		This.lUpsizeDateToDate     = lower(This.GetConfigSetting('UpsizeDateToDate')) = 'true'
		This.lUpsizeFieldsNullable = lower(This.GetConfigSetting('UpsizeFieldsNullable')) = 'true'
		This.lClusteredIndexes     = lower(This.GetConfigSetting('ClusteredIndexes')) = 'true'
		This.lCreateDefaults       = lower(This.GetConfigSetting('CreateDefaults')) = 'true'
		lcBlankDateValue           = This.GetConfigSetting('BlankDateValue')
		This.cBlankDateValue       = iif(lcBlankDateValue = 'NULL', NULL, lcBlankDateValue)
		This.cDefaultDateValue     = This.GetConfigSetting('DefaultDateValue')
	endfunc

*==============================================================================
* Gets a single configuration setting.
*==============================================================================

	function GetConfigSetting(tcTagName, tlFolder)
		local lcBeginDelim, ;
			lcEndDelim, ;
			lcValue
		if empty(This.cConfigSettings)
			This.cConfigSettings = filetostr(This.cConfigFile)
		endif empty(This.cConfigSettings)
		lcBeginDelim = '<' + lower(alltrim(tcTagName)) + '>'
		lcEndDelim   = '</' + lower(alltrim(tcTagName)) + '>'
		lcValue      = strextract(This.cConfigSettings, lcBeginDelim, ;
			lcEndDelim, 1, 1)
		if tlFolder
			lcValue = addbs(fullpath(lcValue, This.cConfigFile))
		endif tlFolder
		return lcValue
	endfunc

*==============================================================================
* Open a SQL Server connection.
*==============================================================================

	function OpenConnection()
		This.DisplayStatus('Opening SQL Server connection')
		This.nHandle = sqlstringconnect(This.cConnectionString)
		if This.nHandle > 0
			sqlsetprop(This.nHandle, 'QueryTimeOut', 0)
		else
			error 1526, This.GetSQLErrorDetails()
		endif This.nHandle > 0
	endfunc

*==============================================================================
* Resets certain folders.
*==============================================================================

	function ResetFolders
		This.DisplayStatus('Resetting folders')
		This.CreateFolder(This.cConvertedDataFolder)
		do case
			case This.cFreeTablesFolder = This.cConvertedDataFolder
				erase (This.cConvertedDataFolder + '*.dbc')
				erase (This.cConvertedDataFolder + '*.dbx')
				erase (This.cConvertedDataFolder + '*.dct')
			otherwise
				erase (This.cConvertedDataFolder + '*.*')
		endcase
		if directory(This.cUpsizeFolder)
			erase (This.cUpsizeFolder + '*.*')
			rd (This.cUpsizeFolder)
		endif directory(This.cUpsizeFolder)
	endfunc

*==============================================================================
* Execute the specified script file in SQL Server.
*==============================================================================

	function ExecuteSQLScriptFile(tcCommandFile)
		local lcSQLCommand
		lcSQLCommand = filetostr(tcCommandFile)
		This.ExecuteSQLCommand(lcSQLCommand)
	endfunc

*==============================================================================
* Execute the specified command in SQL Server.
*==============================================================================

	function ExecuteSQLCommand(tcCommand)
		local lnResult
		lnResult  = sqlexec(This.nHandle, tcCommand)
		if lnResult < 0
			aerror(laError)
			do case
				case 'Changed database context' $ laError[3]
				case This.lRaiseError
					error 1526, This.GetSQLErrorDetails()
				otherwise
					This.DisplayStatus(This.GetSQLErrorDetails())
			endcase
		endif lnResult < 0
	endfunc

*==============================================================================
* Add the free tables to the DBC.
*==============================================================================

	function AddTablesToDBC()
		local lcAlias, ;
			lcTable, ;
			lcKeyColumn, ;
			llRetry, ;
			loException as Exception, ;
			lcAlter, ;
			laFields[1], ;
			lnTotalFields, ;
			lnFieldNo, ;
			lcFieldName, ;
			lcFieldType, ;
			lcCode
		if This.lDebugMode
			set step on 
		endif This.lDebugMode
		select UpsizeTables

* Erase any existing FixDates and FixVarchar scripts.

		erase (This.cTempFolder + 'FixDates.sql')
		erase (This.cTempFolder + 'FixVarchar.sql')

* Process all tables.

		scan
			lcAlias     = trim(TABLE)
			lcTable     = This.cConvertedDataFolder + lcAlias
			lcKeyColumn = trim(KEYCOL)
			This.DisplayStatus('Preparing to add ' + lcAlias + ' to DBC')

* Open the table and clean up the data as necessary.

			use (lcTable) alias (lcAlias) exclusive in 0
			This.ScrubData(lcAlias)
			use in select(lcAlias)

* Add it to the DBC and specify the primary key if there is one.

			This.LogMessage('Adding ' + lcAlias + ' to DBC')
			add table (lcTable)
			if empty(lcKeyColumn)
				use (lcTable) alias (lcAlias) exclusive in 0
			else
				This.LogMessage('Adding ' + lcKeyColumn + ' as primary key')
				alter table (lcAlias) alter column (lcKeyColumn) not null
				alter table (lcAlias) add primary key &lcKeyColumn tag (lcKeyColumn)
			endif empty(lcKeyColumn)

* Execute anything necessary after adding the table to the DBC.

			This.AfterAddTableToDBC(lcAlias)

* Set all logical fields to have .F. as the default value.

			lcAlter       = ''
			lnTotalFields = afields(laFields, lcAlias)
			for lnFieldNo = 1 to lnTotalFields
				lcFieldName = laFields[lnFieldNo, 1]
				lcFieldType = laFields[lnFieldNo, 2]
				if lcFieldType $ 'L'
					lcAlter = lcAlter + ' alter column ' + lcFieldName + ' L default .F.'
				endif lcFieldType $ 'L'
			next lnFieldNo
			if not empty(lcAlter)
				alter table (lcAlias) &lcAlter
			endif not empty(lcAlter)
			use in select(lcAlias)
		endscan

* Clean up backup files.

		erase (This.cConvertedDataFolder + '*.bak')
		erase (This.cConvertedDataFolder + '*.tbk')
		return
	endfunc

*==============================================================================
* Fix the data in the table by replacing bad dates and numeric overflows and
* handling other issues.
*==============================================================================

	function ScrubData(tcAlias)
		local lnSelect, ;
			lnTotalFields, ;
			laFields[1], ;
			ldMinimum, ;
			llFirst, ;
			lcVarcharSQL, ;
			lnFieldNo, ;
			lcFieldName, ;
			lcFieldType, ;
			lcCommand, ;
			lnLength, ;
			lnDecimal, ;
			lnMinimum, ;
			lnMaximum, ;
			lcField, ;
			lcTable, ;
			llUsed, ;
			lnID
		This.LogMessage('Scrubbing data for ' + tcAlias)
		lnSelect = select()
		select (tcAlias)
		lnTotalFields = afields(laFields)
		ldMinimum     = date(1753, 1, 1)
		llFirst       = .T.
		lcVarcharSQL  = ''
		for lnFieldNo = 1 to lnTotalFields
			lcFieldName = laFields[lnFieldNo, 1]
			lcFieldType = laFields[lnFieldNo, 2]
			do case

* Replace bad dates with empty. If we're using SQL Bulk Copy, create a script
* to replace empty dates (which are upsized as 1899) with null.

				case lcFieldType $ 'DT'
					lcCommand = [REPLACE ] + lcFieldName + ;
						[ WITH {//} FOR ] + lcFieldName + [ > {//} AND ] + ;
						lcFieldName + [ < ldMinimum]
					&lcCommand
					if This.lUseSQLBulkCopy
						text to lcSQL noshow textmerge pretext 2
						update [<<tcAlias>>] set [<<lcFieldName>>] = null where [<<lcFieldName>>] = '1899-12-30'
						
						endtext
						strtofile(lcSQL, This.cTempFolder + 'FixDates.sql', .T.)
					endif This.lUseSQLBulkCopy

* Handle integer fields. Since this may be an ID field, we need to log which
* record it's in. Also looking for "*" in the field works more reliably than
* the code used for YN fields.

				case lcFieldType = 'I'
					scan for '*' $ transform(&lcFieldName)
						This.LogMessage('*** Table ' + tcAlias + ', record ' + ;
							transform(recno()) + ': bad value in ' + ;
							lcFieldName)
						replace &lcFieldName with 2147483647
					endscan

* Handle numeric fields.

				case lcFieldType $ 'YN'
					lnLength  = laFields[lnFieldNo, 3]
					lnDecimal = laFields[lnFieldNo, 4]
					set decimals to lnDecimal
					if lnDecimal = 0
						lnMinimum = val('-' + replicate('9', lnLength - 1))
						lnMaximum = val(replicate('9', lnLength))
					else
						lnMinimum = val('-' + ;
							replicate('9', lnLength - lnDecimal - 2) + '.' + ;
							replicate('9', lnDecimal))
						lnMaximum = val(replicate('9', lnLength - lnDecimal - 1) + ;
							'.' + replicate('9', lnDecimal))
					endif

					lcCommand = [REPLACE ] + lcFieldName + ;
						[ WITH 0 FOR ] + lcFieldName + [ < lnMinimum]
					&lcCommand

					lcCommand = [REPLACE ] + lcFieldName + ;
						[ WITH 0 FOR ] + lcFieldName + ;
						[ > lnMaximum OR '*' $ TRANSFORM(] + lcFieldName + [)]
					&lcCommand

* SQL Bulk Copy doesn't trim character fields when upsizing to varchar, so create
* a script to do that if we're supposed to.

				case lcFieldType = 'C' and This.lUseSQLBulkCopy and ;
					This.lUpsizeCharToVarchar
					if llFirst
						text to lcVarcharSQL noshow textmerge pretext 2
						update [<<tcAlias>>] set [<<lcFieldName>>] = RTRIM([<<lcFieldName>>])
						endtext
						llFirst = .F.
					else
						text to lcVarcharSQL additive noshow textmerge pretext 2
						, [<<lcFieldName>>] = RTRIM([<<lcFieldName>>])
						endtext
					endif llFirst
			endcase
		next lnFieldNo

* Update the FixVarchar script if necessary.

		if not empty(lcVarcharSQL)
			strtofile(lcVarcharSQL + ccCRLF, ;
				This.cTempFolder + 'FixVarchar.sql', .T.)
		endif not empty(lcVarcharSQL)
		select (lnSelect)
		return
	endfunc

*==============================================================================
* Execute anything necessary after adding the table to the DBC (abstract in
* this class).
*==============================================================================

	function AfterAddTableToDBC(tcAlias)
	endfunc

*==============================================================================
* Setup the Upsizing Wizard.
*==============================================================================

	function SetupUpsizingWizard
		local lcLibrary

* Create an UpsizeEngine object from UpsizingWizard.app and set its properties.

		set procedure to (This.cApplicationFolder + 'UpsizingWizard.app') ;
			additive
		lcLibrary    = 'Wizusz.prg'
		This.oEngine = newobject('UpsizeEngine', lcLibrary)
		This.SetWizardSettings()

* Do the table analysis.

		This.oEngine.AnalyzeTables()
		This.oEngine.SelectAllTables()
	endfunc

*==============================================================================
* Set the Upsizing Wizard settings. Some of these are set to their default
* values just so they appear here as a reminder of what they're set to.
*==============================================================================

	function SetWizardSettings
		with This.oEngine
			.lQuiet               = This.lQuiet
			.MasterConnHand       = This.nHandle
			.DisconnectOnExit     = .F.
			.ServerDBName         = This.cSQLDatabaseName
			.SourceDB             = This.cDBCFileName
			.CreateNewDB          = This.lCreateDatabase
			.DoUpsize             = .T.
			.DoReport             = This.lCreateReport
			.DoScripts            = .F.
			.DropLocalTables      = .F.
			.DropExistingDatabase = This.lDropExistingDatabase
			.ExportClustered      = This.lClusteredIndexes
			.ExportDefaults       = .T.
				&& although they don't exist in free tables, we'll add them
				&& for logical columns
			.CreateDefaults       = This.lCreateDefaults
			.ExportDRI            = .F.
				&& don't exist in free tables
			.ExportIndexes        = .T.
			.ExportRelations      = .F.
				&& don't exist in free tables
			.ExportSavePwd        = .F.
			.ExportStructureOnly  = .F.
			.ExportTableToView    = .F.
				&& don't exist in free tables
			.ExportViewToRmt      = .F.
				&& don't exist in free tables
			.ExportValidation     = .F.
				&& don't exist in free tables
			.ExportComments       = .F.
				&& don't exist in free tables
			.BlankDateValue       = This.cBlankDateValue
			.DefaultDateValue     = This.cDefaultDateValue
			.NewDir               = This.cLogFolder
			.NormalShutdown       = .T.
			.NotUseBulkInsert     = .F.
			.NullOverride         = 2
				&& General and Memo. Notice we don't set it to 3 (all fields) if
				&& This.lUpsizeFieldsNullable is .T.; if we did that, it would
				&& undo what SetupUpsizingFields does for handling that setting
			.OverWrite            = .T.
			.UserUpsizeMethod     = 6
				&& Fast Export if Bulk Insert Fails
		endwith
	endfunc

*==============================================================================
* Setup the fields to be processed by the Upsizing Wizard.
*==============================================================================

	function SetupUpsizingFields

* Do the field analysis.

		This.oEngine.AnalyzeFields()

* Debug if we're supposed to.

		if This.lDebugMode
			set step on 
		endif This.lDebugMode

* Use varchar instead of char and date instead of datetime for D fields if
* we're supposed to.

		select (This.oEngine.EnumFieldsTbl)
		if This.lUpsizeCharToVarchar
			replace RMTTYPE with 'varchar' for RMTTYPE  = 'char'
		endif This.lUpsizeCharToVarchar
		if This.lUpsizeDateToDate
			replace RMTTYPE with 'date' for DATATYPE = 'D'
		endif This.lUpsizeDateToDate

* Make all fields nullable except PK, non-nullable FK, and logical.

		if This.lUpsizeFieldsNullable
			replace all RMTNULL with .T. for DATATYPE <> 'L'
			select UpsizeTables
			scan for not empty(KeyCol)
				select (This.oEngine.EnumFieldsTbl)
				locate for TBLNAME = lower(trim(UpsizeTables.Table)) and ;
					FLDNAME = lower(trim(UpsizeTables.KeyCol))
				replace RMTNULL with .F.
			endscan for not empty(KeyCol)
			if used('UpsizeRelations')
				select UpsizeRelations
				scan
					select (This.oEngine.EnumFieldsTbl)
					locate for TBLNAME = lower(trim(UpsizeRelations.Parent)) and ;
						FLDNAME = lower(trim(UpsizeRelations.ParentKey))
					replace RMTNULL with .F.
					if not UpsizeRelations.Nullable
						locate for TBLNAME = lower(trim(UpsizeRelations.Child)) and ;
							FLDNAME = lower(trim(UpsizeRelations.ChildKey))
						replace RMTNULL with .F.
					endif not UpsizeRelations.Nullable
				endscan
			endif used('UpsizeRelations')
		endif This.lUpsizeFieldsNullable
	endfunc

*==============================================================================
* Upsize the tables.
*==============================================================================

	function Upsize
		local lnSelect, ;
			lcTable, ;
			llReturn

* Set up the progress events.

		bindevent(This.oEngine, 'InitProcess',   This, 'InitProcess')
		bindevent(This.oEngine, 'UpdateProcess', This, 'UpdateProcess')

* If we're using SQL Bulk Copy, we'll override the SendData method by using
* ourselves as the Upsizing Wizard extension object.

		if This.lUseSQLBulkCopy
			This.oEngine.oExtension = This
		endif This.lUseSQLBulkCopy

* Do the upsizing.

		This.oEngine.ProcessOutput()

* Reset the query timeout because the Upsizing Wizard changes it.

		sqlsetprop(This.nHandle, 'QueryTimeOut', 0)

* Perform any post-upsize database changes.

		lnSelect = select()
		select (This.oEngine.EnumTablesTbl)
		scan for EXPORTED
			lcTable = trim(TBLNAME)
			select UpsizeTables
			seek upper(lcTable)
			if not empty(POSTUPSIZE)
				This.DisplayStatus('Performing post-upsize for ' + lcTable)
				try
					This.ExecuteSQLCommand(POSTUPSIZE)
				catch to loException
					This.DisplayStatus('*** Error #' + transform(loException.ErrorNo) + ;
						' occurred in post-upsize code: ' + loException.Message)
				endtry
			endif not empty(POSTUPSIZE)
		endscan for EXPORTED
		select (lnSelect)

* Wind up.

		unbindevents(This.oEngine)
		llReturn = not This.oEngine.HadError
		if not llReturn
			This.DisplayStatus('*** Error in UpsizingWizard: ' + ;
				This.oEngine.ErrorMessage)
		endif not llReturn
		This.oEngine = NULL
		return llReturn
	endfunc

*==============================================================================
* Use SQLBulkCopy to perform the upsizing.
*==============================================================================

	function SendData(toEngine)
		local loBridge, ;
			loBulkCopy, ;
			lcServer, ;
			lcUser, ;
			lcPassword, ;
			lcTable
		do wwDotNetBridge
		loBridge = GetwwDotNetBridge()
		loBridge.LoadAssembly('SQLBulkCopy.dll')
		loBulkCopy = loBridge.CreateInstance('SQLBulkCopy.SQLBulkCopy')
		loBulkCopy.SourceConnectionString = 'provider=vfpoledb;Data Source=' + ;
			This.cConvertedDataFolder
		lcServer   = strextract(This.cConnectionString, 'server=', ';', 1, 1 + 2)
		lcUser     = strextract(This.cConnectionString, 'uid=',    ';', 1, 1 + 2)
		lcPassword = strextract(This.cConnectionString, 'pwd=',    ';', 1, 1 + 2)
		loBulkCopy.DestinationConnectionString = 'Data Source=' + lcServer + ;
			';Initial Catalog=' + This.cSQLDatabaseName + iif(empty(lcUser), '', ';user=' + lcUser) + ;
			iif(empty(lcPassword), '', ';password=' + lcPassword) + ;
			iif('trusted_connection=yes' $ lower(This.cConnectionString), ';Integrated Security=true', '')
		close databases
			&& we have to close it because we opened it exclusive
		select UpsizeTables
		scan
			lcTable = trim(TABLE)
			This.DisplayStatus('Upsizing data for ' + lcTable)
			try
				loBulkCopy.LoadData(lcTable, 'dbo.[' + lcTable + ']')
			catch to loException
				This.DisplayStatus('*** Error #' + transform(loException.ErrorNo) + ;
					' occurred in line ' + transform(loException.LineNo) + ;
					' of ' + loException.Procedure + ;
					': ' + loException.Message)
			endtry
		endscan
		open database (This.cDBCFileName) exclusive
		return .F.
	endproc

*==============================================================================
* Execute any custom SQL scripts.
*==============================================================================

	function ExecuteCustomSQLScripts
		local lnFiles, ;
			laFiles[1], ;
			lnI, ;
			lcScript
		This.DisplayStatus('Executing custom SQL scripts')
		lnFiles = adir(laFiles, This.cCustomSQLScripts + '*.sql')
		for lnI = 1 to lnFiles
			lcScript = laFiles[lnI, 1]
			This.DisplayStatus('Executing ' + lcScript)
			This.ExecuteSQLScriptFile(This.cCustomSQLScripts + lcScript)
		next lnI
	endfunc

*==============================================================================
* Create foreign keys constraints.
*==============================================================================

	function CreateFKConstraints
		local lcCreateTemplate, ;
			lnSelect, ;
			lcParent, ;
			lcParentKey, ;
			lcChild, ;
			lcChildKey, ;
			lcFK, ;
			lcSQL
		lcCreateTemplate = 'ALTER TABLE [dbo].[<<lcChild>>] WITH NOCHECK ' + ;
			'ADD CONSTRAINT [<<lcFK>>] FOREIGN KEY ([<<lcChildKey>>]) ' + ;
			'REFERENCES [dbo].[<<lcParent>>] ([<<lcParentKey>>])' + ccCRLF + ;
			'ALTER TABLE [dbo].[<<lcChild>>] CHECK CONSTRAINT [<<lcFK>>]' + ;
			ccCRLF + ccCRLF
		lnSelect         = select()
		select UpsizeRelations
		scan
			lcParent    = lower(trim(PARENT))
			lcParentKey = lower(trim(PARENTKEY))
			lcChild     = lower(trim(CHILD))
			lcChildKey  = lower(trim(CHILDKEY))
			lcFK        = 'FK_' + lcChild + '_' + lcParent + ;
				iif(lcParentKey == lcChildKey, '', '_' + lcChildKey)
			lcSQL       = textmerge(lcCreateTemplate)
			This.DisplayStatus('Creating FK constraints for ' + lcParent + ;
				' - ' + lcChild)
			This.ExecuteSQLCommand(lcSQL)
		endscan
		select (lnSelect)
	endfunc

*==============================================================================
* If we're using SQL Bulk Copy, set empty dates to null and trim varchars.
*==============================================================================

	function DataCleanup
		if This.lUseSQLBulkCopy
			if file(This.cTempFolder + 'FixDates.sql')
				This.DisplayStatus('Fixing empty dates')
				This.ExecuteSQLScriptFile(This.cTempFolder + 'FixDates.sql')
			endif file(This.cTempFolder + 'FixDates.sql')
			if file(This.cTempFolder + 'FixVarchar.sql')
				This.DisplayStatus('Removed trailing spaces in varchar')
				This.ExecuteSQLScriptFile(This.cTempFolder + 'FixVarchar.sql')
			endif file(This.cTempFolder + 'FixVarchar.sql')
		endif This.lUseSQLBulkCopy
	endfunc

*==============================================================================
* Handle an error sending a command to SQL Server.
*==============================================================================

	function GetSQLErrorDetails
		local lcDetails, ;
			laError[1]
		lcDetails = ''
		try
			aerror(laError)
			text to lcDetails noshow textmerge flags 1 + 2 pretext 1 + 2 + 4
			Error Message: <<laError[2]>>
			ODBC Error: <<nvl(laError[3], '')>>
			SQL State: <<nvl(laError[4], '')>>
			ODBC Error No: <<transform(nvl(laError[5], 0))>>
			endtext
		catch
			lcDetails = 'GetSQLErrorDetails() Error'
		endtry
		return lcDetails
	endfunc

*==============================================================================
* Helper methods.
*==============================================================================

	function ShellExecute(hwnd, lpVerb, lpFile, lpParameters, lpDirectory, nShowCmd)
		declare integer ShellExecute in SHELL32.dll integer hwnd, string lpVerb, ;
			string lpFile, string lpParameters, string lpDirectory, long nShowCmd
		return ShellExecute(hwnd, lpVerb, lpFile, lpParameters, lpDirectory, nShowCmd)
	endfunc

	function Sleep(dwMilliseconds)
		declare Sleep in KERNEL32.dll integer dwMilliseconds
		return Sleep(dwMilliseconds)
	endfunc

*==============================================================================
* Upsizing logging methods.
*==============================================================================

	function InitProcess(tcTitle, tnBasis)
		local ltNow
		ltNow = datetime()
		if tcTitle <> This.cTitle or ltNow >= This.tNextInit
			This.tNextInit = ltNow + 5
			This.cTitle    = tcTitle
			This.DisplayStatus('Upsizing Data - ' + tcTitle + ;
				iif(tnBasis > 0, ' - ' + transform(tnBasis), ''))
		endif tcTitle <> This.cTitle ...
		return
	endfunc

	function UpdateProcess(tnProgress, tcTask)
		local ltNow
		ltNow = datetime()
		if tcTask <> This.cTask or ltNow >= This.tNextUpdate
			This.tNextUpdate = ltNow + 5
			This.cTask       = tcTask
			This.DisplayStatus('Upsizing Data - ' + tcTask + ;
				iif(tnProgress > 0, ' - ' + transform(tnProgress), ''))
		endif tcTask <> This.cTask ...
		return
	endfunc

*==============================================================================
* Create the specified folder.
*==============================================================================

	function CreateFolder(tcFolder)
		local lcFullPath, ;
			lnTotalLevels, ;
			lnLevelNo, ;
			lcFolder
		lcFullPath = addbs(tcFolder)
		if directory(lcFullPath, 1)
			return .T.
		endif directory(lcFullPath, 1)
		lnTotalLevels = occurs('\', lcFullPath)
		for lnLevelNo = 1 to lnTotalLevels
			lcFolder = left(lcFullPath, at('\', lcFullPath, lnLevelNo))
			if not directory(lcFolder, 1)
				md (lcFolder)
			endif not directory(lcFolder, 1)
		next lnLevelNo
		return directory(lcFullPath, 1)
	endfunc

*==============================================================================
* Displays and logs the current status.
*==============================================================================

	function DisplayStatus(tcMessage)
		if not This.lQuiet
			wait window left(tcMessage + '...', 254) nowait noclear
		endif not This.lQuiet
		if This.lShowProgress
			? tcMessage
		endif This.lShowProgress
		This.LogMessage(tcMessage)
	endfunc

*==============================================================================
* Logs the current status.
*==============================================================================

	function LogMessage(tcMessage)
		local lnNow, ;
			lcTotalTime, ;
			lcElapsedTime
		lnNow = seconds()
		if This.nStartLog = 0
			This.nStartLog = lnNow
			lcTotalTime    = ''
		else
			if lnNow < This.nStartLog
				lnNow = lnNow + 24 * 60 * 60
					&& adjust for past midnight
			endif lnNow < This.nStartLog
			lcTotalTime = ', ' + ;
				transform(lnNow - This.nStartLog, '99,999.999') + ' total'
		endif This.nStartLog = 0
		if This.nLastLog = 0
			lcElapsedTime = ''
		else
			lcElapsedTime = ', ' + ;
				transform(lnNow - This.nLastLog, '99,999.999') + ' elapsed'
		endif This.nLastLog = 0
		strtofile(transform(datetime()) + lcElapsedTime + lcTotalTime + ;
			':   ' + tcMessage + ccCRLF, This.cLogFile, .T.)
		This.nLastLog = lnNow
	endfunc
enddefine
