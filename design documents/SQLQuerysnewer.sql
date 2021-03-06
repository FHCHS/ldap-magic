DROP TABLE FHC_LDAP_sqlusersTable
DROP TABLE FHC_LDAP_gmailusersTable
{CREATE TABLE FHC_LDAP_sqlusersTable(sAMAccountName VarChar(350), givenName VarChar(350), sn VarChar(350), middleName VarChar(350), comment VarChar(350))}
{CREATE TABLE FHC_LDAP_gmailusersTable(sAMAccountName VarChar(350), givenName VarChar(350), sn VarChar(350))}
SELECT DISTINCT uptoDate.* FROM FHC_LDAP_sqlusersTable uptoDate LEFT OUTER JOIN FHC_LDAP_gmailusersTable outofDate ON outofDate.sAMAccountName = uptoDate.sAMAccountName WHERE outofDate.sAMAccountName IS NULL
SELECT DISTINCT FHC_LDAP_sqlusersTable.givenName, FHC_LDAP_sqlusersTable.sn, FHC_LDAP_sqlusersTable.sAMAccountName, FHC_LDAP_sqlusersTable.middleName FROM FHC_LDAP_sqlusersTable INNER JOIN FHC_LDAP_gmailusersTable ON FHC_LDAP_sqlusersTable.sAMAccountName = FHC_LDAP_gmailusersTable.sAMAccountName AND (FHC_LDAP_gmailusersTable.givenName + FHC_LDAP_gmailusersTable.sn + FHC_LDAP_gmailusersTable.sAMAccountName ) <> (FHC_LDAP_sqlusersTable.givenName + FHC_LDAP_sqlusersTable.sn + FHC_LDAP_sqlusersTable.sAMAccountName ) WHERE FHC_LDAP_gmailusersTable.givenName <> '' OR FHC_LDAP_gmailusersTable.sn <> '' 
DROP TABLE FHC_LDAP_gmailNicknamesTable
DROP TABLE FHC_LDAP_loginsWONicknamesTable
DROP TABLE FHC_LDAP_adNicknamesTable
DROP TABLE FHC_LDAP_sqlNicknamesTable
DROP TABLE FHC_LDAP_nicknamesToUpdateDB
DROP TABLE FHC_LDAP_nicknamesFilteredDuplicates
SELECT DISTINCT RTRIM(soc_sec) AS soc_sec, RTRIM(e_mail) AS e_mail INTO FHC_LDAP_sqlNicknamesTable FROM address WHERE preferred = 1
{CREATE TABLE FHC_LDAP_adNicknamesTable(sAMAccountName VarChar(350), givenName VarChar(350), middleName VarChar(350), sn VarChar(350), mail VarChar(350))}
{CREATE TABLE FHC_LDAP_gmailNicknamesTable(soc_sec VarChar(350), nickname VarChar(350), Email VarChar(350))}
DROP TABLE FHC_LDAP_gmailusersTable
{CREATE TABLE FHC_LDAP_gmailusersTable(sAMAccountName VarChar(350), givenName VarChar(350), sn VarChar(350))}
SELECT DISTINCT uptoDate.* INTO FHC_LDAP_loginsWONicknamesTable FROM FHC_LDAP_gmailusersTable uptoDate LEFT OUTER JOIN FHC_LDAP_gmailNicknamesTable outofDate ON outofDate.soc_sec = uptoDate.sAMAccountName WHERE outofDate.soc_sec IS NULL
SELECT DISTINCT FHC_LDAP_sqlusersTable.* FROM FHC_LDAP_sqlusersTable INNER JOIN FHC_LDAP_loginsWONicknamesTable ON FHC_LDAP_sqlusersTable.sAMAccountName = FHC_LDAP_loginsWONicknamesTable.sAMAccountName
DROP TABLE FHC_LDAP_gmailNicknamesTable
{CREATE TABLE FHC_LDAP_gmailNicknamesTable(soc_sec VarChar(350), nickname VarChar(350), Email VarChar(350))}
SELECT DISTINCT extTable.soc_sec, (SELECT TOP 1 FHC_LDAP_gmailNicknamesTable.Email FROM FHC_LDAP_gmailNicknamesTable LEFT JOIN FHC_LDAP_sqlusersTable ON FHC_LDAP_gmailNicknamesTable.soc_sec = FHC_LDAP_sqlusersTable.sAMAccountName WHERE FHC_LDAP_gmailNicknamesTable.soc_sec = extTable.soc_sec ORDER BY FHC_LDAP_sqlusersTable.sAMAccountName, CHARINDEX(FHC_LDAP_sqlusersTable.sn, FHC_LDAP_gmailNicknamesTable.nickname) DESC, CHARINDEX(FHC_LDAP_sqlusersTable.givenName, FHC_LDAP_gmailNicknamesTable.nickname) DESC, CHARINDEX(FHC_LDAP_sqlusersTable.middleName, FHC_LDAP_gmailNicknamesTable.nickname) DESC) AS Email INTO FHC_LDAP_nicknamesFilteredDuplicates FROM FHC_LDAP_gmailNicknamesTable AS extTable ORDER BY extTable.soc_sec
SELECT DISTINCT FHC_LDAP_nicknamesFilteredDuplicates.soc_sec, FHC_LDAP_nicknamesFilteredDuplicates.Email INTO FHC_LDAP_nicknamesToUpdateDB FROM FHC_LDAP_nicknamesFilteredDuplicates INNER JOIN FHC_LDAP_sqlNicknamesTable ON FHC_LDAP_nicknamesFilteredDuplicates.soc_sec = FHC_LDAP_sqlNicknamesTable.soc_sec WHERE FHC_LDAP_nicknamesFilteredDuplicates.Email NOT IN ( SELECT FHC_LDAP_sqlNicknamesTable.e_mail FROM FHC_LDAP_sqlNicknamesTable )
UPDATE address SET address.e_mail2 = address.e_mail FROM address INNER JOIN FHC_LDAP_nicknamesToUpdateDB ON address.soc_sec = FHC_LDAP_nicknamesToUpdateDB.soc_sec WHERE address.e_mail not like '%students.fhchs.edu%' AND address.e_mail <> '' AND address.e_mail <> '?' AND address.e_mail IS NOT NULL AND preferred = 1
UPDATE address SET address.e_mail = FHC_LDAP_nicknamesToUpdateDB.Email FROM address INNER JOIN FHC_LDAP_nicknamesToUpdateDB ON address.soc_sec = FHC_LDAP_nicknamesToUpdateDB.soc_sec WHERE preferred = 1
SELECT DISTINCT FHC_LDAP_nicknamesFilteredDuplicates.soc_sec, FHC_LDAP_nicknamesFilteredDuplicates.Email FROM FHC_LDAP_nicknamesFilteredDuplicates INNER JOIN FHC_LDAP_adNicknamesTable ON FHC_LDAP_nicknamesFilteredDuplicates.soc_sec = FHC_LDAP_adNicknamesTable.sAMAccountName WHERE (FHC_LDAP_adNicknamesTable.sAMAccountName + FHC_LDAP_adNicknamesTable.mail ) <> (FHC_LDAP_nicknamesFilteredDuplicates.soc_sec + FHC_LDAP_nicknamesFilteredDuplicates.Email )
SELECT DISTINCT FHC_LDAP_sqlusersTable.*, FHC_LDAP_nicknamesFilteredDuplicates.Email FROM FHC_LDAP_sqlusersTable INNER JOIN FHC_LDAP_nicknamesFilteredDuplicates ON FHC_LDAP_sqlusersTable.sAMAccountName = FHC_LDAP_nicknamesFilteredDuplicates.soc_sec
