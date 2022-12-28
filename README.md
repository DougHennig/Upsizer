# Upsizer

The [VFPX Upsizing Wizard](https://github.com/VFPX/upsizingwizard) allows you to upsize the tables in a database container to a SQL Server database. It can be used visually by running UpsizingWizard.app or programmatically; see TestEngine.prg that comes with the Upsizing Wizard for an example.

One thing the Upsizing Wizard doesn't do is support free tables. The Upsizer utility gives the Upsizing Wizard the ability to do that.

## How it Works

Upsizer starts by copying the VFP free tables you want to upsize to a new folder. It does this because it makes changes to the tables, including adding them to a DBC because Upsizing Wizard requires that, and we don't want to make changes to the original tables.

It then creates a DBC named Upsizer.dbc in the copied tables folder and adds the free tables to it. As it does that, it also performs some clean up on the tables:

* It replaces bad dates (those prior to 1753-01-01) with empty dates.

* It replaces bad numeric values (those appearing as "****" in a BROWSE window) with 0.

* If you've specified a primary key field for a table, it makes that field non-nullable and creates a primary key index on it.

* It sets .F. as the default value for all logical fields.

Upsizer then instantiates the Upsizing Wizard engine object and tells it to upsize the DBC and its tables. Once that's done, it does some cleanup of the SQL Server tables, trimming Varchar values and converting 1899-12-30 date values (the values inserted into SQL Server for VFP blank dates) to null.

Upsizer logs every thing it does to a log file which you can review for problems afterward.

## Upsizer Files

Upsizer consists of the following files:

* Upsizer.prg: the main program.

* Upsizer.config: the configuration file for Upsizer. This file must be edited to specify the correct settings; see the [Configuring Upsizer](#configuring) section.

* UpsizingWizard.app: the [VFPX Upsizing Wizard](https://github.com/VFPX/upsizingwizard) used to perform the actual upsizing.

* SQLBulkCopy.dll: a .NET DLL that provides [fast SQL bulk copy](https://github.com/DougHennig/SQLBulkCopy) functionality.

* ClrHost.dll, wwDotNetBridge.dll, and wwDotNetBridge.prg: Rick Strahl's [wwDotNetBridge](https://github.com/RickStrahl/wwDotnetBridge), needed to use SQLBulkCopy.dll.

## <a name="configuring">Configuring Upsizer</a>

Upsizer.config specifies the settings used by Upsizer, such as the connection string for the SQL Server database. Folder paths can be relative to the folder containing Upsizer; for example, ```<ConvertedDataFolder>ConvertedData</ConvertedDataFolder>``` means copy the source VFP tables to the ConvertedData subdirectory of the application folder.

Here are the settings specified in Upsizer.config:

```xml
<?xml version = "1.0" encoding="Windows-1252" standalone="yes"?>
<VFPData>
   <Upsizer>
      <ConnectionString>The SQL Server connection string</ConnectionString>
      <SQLDatabaseName>The name of the SQL Server database</SQLDatabaseName>
      <CreateDatabase>true to create the SQL Server database; false if it already exists</CreateDatabase>
      <DropExistingDatabase>true to drop the SQL Server database if it already exists; false to drop the tables in the database but not drop the database</DropExistingDatabase>
      <FreeTablesFolder>The folder where the VFP tables are copied from (used if CopySourceTables is true)</FreeTablesFolder>
      <CopySourceTables>true to copy the VFP tables from the folder specified in FreeTablesFolder to the one specified in ConvertedDataFolder; false to use an existing set of tables in ConvertedDataFolder</CopySourceTables>
      <ConvertedDataFolder>The folder to copy the VFP tables to or the location of the tables to upsize if CopySourceTables is false</ConvertedDataFolder>
      <CustomSQLScriptsFolder>The folder containing any SQL scripts to execute after upsizing</CustomSQLScriptsFolder>
      <LogFolder>The location for log files</LogFolder>
      <TimestampLogFile>true to timestamp the log file</TimestampLogFile>
      <QuietMode>false to display messages, true to not</QuietMode>
      <DebugMode>true to SET STEP ON when running from the IDE</DebugMode>
      <SQLBulkCopy>true to use SQLBulkCopy; false to use bulk XML load</SQLBulkCopy>
      <ShowProgress>true to display progress</ShowProgress>
      <CreateReport>true to have the Upsizing Wizard create tables and reports showing what was upsized and what failed</CreateReport>
      <UpsizeCharToVarchar>true to upsize Character fields to Varchar; false to upsize them as Char</UpsizeCharToVarchar>
      <UpsizeDateToDate>true to upsize Date fields to Date; false to upsize them as DateTime</UpsizeDateToDate>
      <UpsizeFieldsNullable>true to mark all fields as accepting null values; false to use the setting of the VFP fields</UpsizeFieldsNullable>
      <ClusteredIndexes>true to upsize primary keys as clustered indexes; false to upsize them as regular indexes</ClusteredIndexes>
      <CreateDefaults>true to create default values for every field</CreateDefaults>
      <BlankDateValue>the value to use for blank date fields. such as NULL or a specific date</BlankDateValue>
      <DefaultDateValue>the default value to use for Date and DateTime fields if CreateDefaults is true; "DATE()" is automatically converted to the SQL Server equivalent "GETDATE()"</DefaultDateValue>
   </Upsizer>
</VFPData>
```

## Running Upsizer

To upsize a set of VFP free tables, edit Upsizer.config (see the previous section) as necessary, then run Upsizer.prg. It displays progress messages if the QuietMode and ShowProgress settings in Upsizer.config are false.

If you're running Upsizer multiple times without making any changes to the data, such as when testing, you can save a bit of time by setting the CopySourceTables setting to false; in that case, it assumes the folder specified in the ConvertedDataFolder setting contains the tables to upsize and they've already been added to a DBC.

After the upsizing is completed, any .SQL files in the folder specified in the CustomSQLScriptsFolder setting are executed. These scripts can contain any T-SQL commands you wish, such as creating stored procedures, functions, views, and so on.

## Metadata Tables

There are two metadata tables used by Upsizer:

* UpsizeTables: contains the name of each table to be upsized, the name of its primary key field, and any post-upsize SQL script to execute (the latter can contain any T-SQL commands you wish to execute after that table has been upsized).

* UpsizeRelations: specifies the relationships between tables. This table has columns for the names of the parent table, its primary key field, the child table, the foreign key field, and whether the foreign key field can accept null values or not.

These tables are optional: if they don't exist, all tables in the specified free tables folder are upsized, none have primary keys assigned, and no foreign key constraints are created.

You can add records to UpsizeTables.dbf and UpsizeRelations.dbf manually if you wish, or you can DO FORM TableBuilder.scx in the TableBuilder folder.

## Subclassing Upsizer

Although subclassing isn't necessary in many cases because of the configuration settings specified in Upsizer.config, the Upsizer class in Upsizer.prg was designed to be subclassed if necessary.

The likely places to override in subclass are:

### ScrubData
This method, which is called if CopySourceTables is true, cleans up data in the specified table so it upsizes better (for example, converting illegal date and numeric values). In a subclass, you could do additional clean up, even table-specific. For example:

```foxpro
function ScrubData(tcAlias)
dodefault(tcAlias)
do case
	case tcAlias = 'SOMETABLE'
* do special data processing e.g. fix bad data,
* remove duplicate records, fix FKs (0 should be null), etc.
endcase
```

### SetupUpsizingFields
This method adjusts the cursor created and used by the Upsizing Wizard to customize how fields are upsized. For example, if the UpsizeCharToVarchar setting is true, it uses this code to do so:

```foxpro
replace RMTTYPE with 'varchar' for RMTTYPE  = 'char'
```

In a subclass, DODEFAULT() and then make any desired changes to cursor specified in This.oEngine.EnumFieldsTbl.

### SetWizardSettings
This method sets the settings of the Upsizing Wizard engine object as necessary for Upsizer and as defined by the settings in Upsizer.config. In a subclass, DODEFAULT() and then make any desired changes to This.oEngine settings.

### AfterAddTableToDBC
This method, which is abstract in Upsizer.prg, allows you to make any changes to the specified table (passed as a parameter) after adding it to DBC.

## Helping with this project

See [How to contribute to Upsizer](.github/CONTRIBUTING.md) for details on how to help with this project.

## Releases

### 2022-12-28

Initial release.
