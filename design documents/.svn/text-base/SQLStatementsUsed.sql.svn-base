DROP TABLE FHC_TEST_sqlusersTable
DROP TABLE FHC_TEST_gmailusersTable

CREATE TABLE FHC_TEST_gmailusersTable --missed command
CREATE TABLE FHC_TEST_gmailusersTable(sAMAccountName VarChar(350), givenName VarChar(350), sn VarChar(350))

--check what needs to be added to gmail
SELECT DISTINCT uptoDate.* FROM FHC_TEST_sqlusersTable uptoDate LEFT OUTER JOIN FHC_TEST_gmailusersTable outofDate ON outofDate.sAMAccountName = uptoDate.sAMAccountName WHERE outofDate.sAMAccountName IS NULL;

--check which gmail accounts need updated fnames lnames
SELECT DISTINCT FHC_TEST_sqlusersTable.givenName, FHC_TEST_sqlusersTable.sn, FHC_TEST_sqlusersTable.sAMAccountName, FHC_TEST_sqlusersTable.middleName FROM FHC_TEST_sqlusersTable INNER JOIN FHC_TEST_gmailusersTable ON FHC_TEST_sqlusersTable.sAMAccountName = FHC_TEST_gmailusersTable.sAMAccountName AND (FHC_TEST_gmailusersTable.givenName + FHC_TEST_gmailusersTable.sn + FHC_TEST_gmailusersTable.sAMAccountName ) <> (FHC_TEST_sqlusersTable.givenName + FHC_TEST_sqlusersTable.sn + FHC_TEST_sqlusersTable.sAMAccountName ) WHERE FHC_TEST_gmailusersTable.givenName <> '' OR FHC_TEST_gmailusersTable.sn <> '' 

--clean up for nickname portion
DROP TABLE FHC_TEST_gmailNicknamesTable
DROP TABLE FHC_TEST_loginsWONicknamesTable
DROP TABLE FHC_TEST_adNicknamesTable
DROP TABLE FHC_TEST_sqlNicknamesTable
DROP TABLE FHC_TEST_nicknamesToUpdateDB

--get base information from address table of nicknames etc
SELECT DISTINCT RTRIM(soc_sec) AS soc_sec, RTRIM(e_mail) AS e_mail INTO FHC_TEST_sqlNicknamesTable FROM address WHERE preferred = 1

--create table for information from AD nicknames etc
CREATE TABLE FHC_TEST_adNicknamesTable(sAMAccountName VarChar(350), givenName VarChar(350), middleName VarChar(350), sn VarChar(350), mail VarChar(350))

--create table for information from gmail nicknames
CREATE TABLE FHC_TEST_gmailNicknamesTable(soc_sec VarChar(350), nickname VarChar(350), Email VarChar(350))

--gmail users have changed since we may have added new ones drop the table and reget the information
DROP TABLE FHC_TEST_gmailusersTable

--create table for information with the gmail user login info
CREATE TABLE FHC_TEST_gmailusersTable(sAMAccountName VarChar(350), givenName VarChar(350), sn VarChar(350))

--create table with the logins which do not have nicknames yet will contain data from administrative users and people in the nurse anethestist program
SELECT DISTINCT uptoDate.* INTO FHC_TEST_loginsWONicknamesTable FROM FHC_TEST_gmailusersTable uptoDate LEFT OUTER JOIN FHC_TEST_gmailNicknamesTable outofDate ON outofDate.soc_sec = uptoDate.sAMAccountName WHERE outofDate.soc_sec IS NULL

--get the lost nicknames (accounts without nicknames)
SELECT DISTINCT FHC_TEST_sqlusersTable.* FROM FHC_TEST_sqlusersTable INNER JOIN FHC_TEST_loginsWONicknamesTable ON FHC_TEST_sqlusersTable.sAMAccountName = FHC_TEST_loginsWONicknamesTable.sAMAccountName

--Nicknames might have changed repull
DROP TABLE FHC_TEST_gmailNicknamesTable

--create table for information from gmail nicknames
CREATE TABLE FHC_TEST_gmailNicknamesTable(soc_sec VarChar(350), nickname VarChar(350), Email VarChar(350))

--create table for nicknames with the closest matching nickname to the persons real name
SELECT DISTINCT extTable.soc_sec, 
	(SELECT TOP 1 FHC_TEST_gmailNicknamesTable.Email 
	FROM FHC_TEST_gmailNicknamesTable LEFT JOIN FHC_TEST_sqlusersTable 
	ON FHC_TEST_gmailNicknamesTable.soc_sec = FHC_TEST_sqlusersTable.sAMAccountName 
	WHERE FHC_TEST_gmailNicknamesTable.soc_sec = extTable.soc_sec 
	ORDER BY FHC_TEST_sqlusersTable.sAMAccountName, 
	CHARINDEX(FHC_TEST_sqlusersTable.sn, FHC_TEST_gmailNicknamesTable.nickname) DESC, 
	CHARINDEX(FHC_TEST_sqlusersTable.givenName, FHC_TEST_gmailNicknamesTable.nickname) DESC, 
	CHARINDEX(FHC_TEST_sqlusersTable.middleName, FHC_TEST_gmailNicknamesTable.nickname) DESC)
	AS Email 
INTO FHC_TEST_nicknamesFilteredDuplicates 
FROM FHC_TEST_gmailNicknamesTable AS extTable 
ORDER BY extTable.soc_sec



--get emails which need to be updated
SELECT DISTINCT FHC_TEST_nicknamesFilteredDuplicates.soc_sec, FHC_TEST_nicknamesFilteredDuplicates.Email 
--INTO FHC_TEST_nicknamesToUpdateDB 
FROM FHC_TEST_nicknamesFilteredDuplicates INNER JOIN FHC_TEST_sqlNicknamesTable 
ON FHC_TEST_nicknamesFilteredDuplicates.soc_sec = FHC_TEST_sqlNicknamesTable.soc_sec 
WHERE FHC_TEST_nicknamesFilteredDuplicates.Email NOT IN 
	( SELECT FHC_TEST_sqlNicknamesTable.e_mail 
	FROM FHC_TEST_sqlNicknamesTable )


--shift addresses for emails which are new from students domain
UPDATE address SET address.e_mail2 = address.e_mail FROM address INNER JOIN FHC_TEST_nicknamesToUpdateDB ON address.soc_sec = FHC_TEST_nicknamesToUpdateDB.soc_sec WHERE address.e_mail not like '%students.fhchs.edu%' AND preferred = 1
UPDATE address SET address.e_mail2 = address.e_mail FROM address INNER JOIN FHC_TEST_nicknamesToUpdateDB ON address.soc_sec = FHC_TEST_nicknamesToUpdateDB.soc_sec WHERE address.e_mail not like '%students.fhchs.edu%' AND address.e_mail <> '' AND address.e_mail <> '?' AND address.e_mail IS NOT NULL AND preferred = 1
SELECT address.e_mail2, address.e_mail FROM address INNER JOIN FHC_TEST_nicknamesToUpdateDB ON address.soc_sec = FHC_TEST_nicknamesToUpdateDB.soc_sec WHERE address.e_mail not like '%students.fhchs.edu%' AND address.e_mail <> '' AND address.e_mail <> '?' AND address.e_mail IS NOT NULL AND preferred = 1

--actually copy in new emails
UPDATE address SET address.e_mail = FHC_TEST_nicknamesToUpdateDB.Email FROM address INNER JOIN FHC_TEST_nicknamesToUpdateDB ON address.soc_sec = FHC_TEST_nicknamesToUpdateDB.soc_sec WHERE preferred = 1
SELECT address.e_mail, FHC_TEST_nicknamesToUpdateDB.Email FROM address INNER JOIN FHC_TEST_nicknamesToUpdateDB ON address.soc_sec = FHC_TEST_nicknamesToUpdateDB.soc_sec WHERE preferred = 1

--check which emails need updating in AD
SELECT DISTINCT FHC_TEST_nicknamesFilteredDuplicates.soc_sec, FHC_TEST_nicknamesFilteredDuplicates.Email, (FHC_TEST_adNicknamesTable.sAMAccountName + FHC_TEST_adNicknamesTable.mail ) as ADmail, (FHC_TEST_nicknamesFilteredDuplicates.soc_sec + FHC_TEST_nicknamesFilteredDuplicates.Email ) as Gmail
FROM FHC_TEST_nicknamesFilteredDuplicates INNER JOIN FHC_TEST_adNicknamesTable 
ON FHC_TEST_nicknamesFilteredDuplicates.soc_sec = FHC_TEST_adNicknamesTable.sAMAccountName 
WHERE (FHC_TEST_adNicknamesTable.sAMAccountName + FHC_TEST_adNicknamesTable.mail )
 <> (FHC_TEST_nicknamesFilteredDuplicates.soc_sec + FHC_TEST_nicknamesFilteredDuplicates.Email )


--get list of nicknames
SELECT DISTINCT FHC_TEST_sqlusersTable.*, FHC_TEST_nicknamesFilteredDuplicates.Email
 FROM FHC_TEST_sqlusersTable INNER JOIN FHC_TEST_nicknamesFilteredDuplicates 
ON FHC_TEST_sqlusersTable.sAMAccountName = FHC_TEST_nicknamesFilteredDuplicates.soc_sec

--get list of information for creating send as aliases
SELECT DISTINCT FHC_TEST_sqlusersTable.*, FHC_TEST_nicknamesFilteredDuplicates.Email FROM FHC_TEST_sqlusersTable INNER JOIN FHC_TEST_nicknamesFilteredDuplicates ON FHC_TEST_sqlusersTable.sAMAccountName = FHC_TEST_nicknamesFilteredDuplicates.soc_sec

select * from FHC_TEST_nicknamesFilteredDuplicates
