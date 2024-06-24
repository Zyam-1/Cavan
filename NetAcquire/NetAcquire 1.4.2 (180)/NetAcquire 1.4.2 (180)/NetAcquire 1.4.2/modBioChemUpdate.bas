Attribute VB_Name = "modBioChemUpdate"
Option Explicit

Public Sub UpdateBioResults()

58710 EnsureColumnExists "BioResults", "DefIndex", "numeric NOT NULL DEFAULT 0"

End Sub

Public Sub UpdateBioTestDefinitions()

      Dim sql As String

58720 On Error GoTo UpdateBioTestDefinitions_Error

58730 EnsureColumnExists "BioTestDefinitions", "AgeFromText", "nvarchar(50) NOT NULL DEFAULT '0 Days'"
58740 EnsureColumnExists "BioTestDefinitions", "AgeToText", "nvarchar(50) NOT NULL DEFAULT '120 Years'"
58750 EnsureColumnExists "BioTestDefinitions", "PrintRefRange", "tinyint NOT NULL DEFAULT 1"

58760 sql = "UPDATE BioTestDefinitions " & _
            "SET AgeFromText = '1 Month', " & _
            "AgeFromDays = 30 WHERE AgeFromDays = 31"
58770 Cnxn(0).Execute sql

58780 sql = "UPDATE BioTestDefinitions " & _
            "SET AgeFromText = '3 Months', " & _
            "AgeFromDays = 90 WHERE AgeFromDays = 91"
58790 Cnxn(0).Execute sql

58800 sql = "UPDATE BioTestDefinitions " & _
            "SET AgeFromText = '1 Year', " & _
            "AgeFromDays = 365 WHERE AgeFromDays = 366"
58810 Cnxn(0).Execute sql

58820 sql = "UPDATE BioTestDefinitions " & _
            "SET AgeFromText = '2 Years', " & _
            "AgeFromDays = 730 WHERE AgeFromDays = 731"
58830 Cnxn(0).Execute sql

58840 sql = "UPDATE BioTestDefinitions " & _
            "SET AgeFromText = '12 Years', " & _
            "AgeFromDays = 4383 WHERE AgeFromDays = 4381"
58850 Cnxn(0).Execute sql

58860 sql = "UPDATE BioTestDefinitions " & _
            "SET AgeFromText = '50 Years', " & _
            "AgeFromDays = 18262 WHERE AgeFromDays = 18251 " & _
            "OR AgeFromDays = 18263"
58870 Cnxn(0).Execute sql

58880 sql = "UPDATE BioTestDefinitions " & _
            "SET AgeFromText = '60 Years', " & _
            "AgeFromDays = 21900 WHERE AgeFromDays = 21901 " & _
            "OR AgeFromDays = 21916"
58890 Cnxn(0).Execute sql

58900 sql = "UPDATE BioTestDefinitions " & _
            "SET AgeFromText = '70 Years', " & _
            "AgeFromDays = 25550 WHERE AgeFromDays = 25551 " & _
            "OR AgeFromDays = 25569"
58910 Cnxn(0).Execute sql

58920 sql = "UPDATE BioTestDefinitions " & _
            "SET AgeFromText = '80 Years', " & _
            "AgeFromDays = 29200 WHERE AgeFromDays = 29201"
58930 Cnxn(0).Execute sql

58940 sql = "UPDATE BioTestDefinitions " & _
            "SET AgeFromText = '120 Years', " & _
            "AgeFromDays = 43830 WHERE AgeFromDays > = 43800"
58950 Cnxn(0).Execute sql

      '''''''''''''

58960 sql = "UPDATE BioTestDefinitions " & _
            "SET AgeToText = '1 Month', " & _
            "AgeToDays = 30 WHERE AgeToDays = 30"
58970 Cnxn(0).Execute sql

58980 sql = "UPDATE BioTestDefinitions " & _
            "SET AgeToText = '3 Months', " & _
            "AgeToDays = 90 WHERE AgeToDays = 90"
58990 Cnxn(0).Execute sql

59000 sql = "UPDATE BioTestDefinitions " & _
            "SET AgeToText = '1 Year', " & _
            "AgeToDays = 365 WHERE AgeToDays = 365"
59010 Cnxn(0).Execute sql

59020 sql = "UPDATE BioTestDefinitions " & _
            "SET AgeToText = '2 Years', " & _
            "AgeToDays = 730 WHERE AgeToDays = 730"
59030 Cnxn(0).Execute sql

59040 sql = "UPDATE BioTestDefinitions " & _
            "SET AgeToText = '50 Years', " & _
            "AgeToDays = 18262 WHERE AgeToDays = 18250 " & _
            "OR AgeToDays = 18262"
59050 Cnxn(0).Execute sql

59060 sql = "UPDATE BioTestDefinitions " & _
            "SET AgeToText = '60 Years', " & _
            "AgeToDays = 21900 WHERE AgeToDays = 21900 " & _
            "OR AgeToDays = 21915"
59070 Cnxn(0).Execute sql

59080 sql = "UPDATE BioTestDefinitions " & _
            "SET AgeToText = '70 Years', " & _
            "AgeToDays = 25550 WHERE AgeToDays = 25550 " & _
            "OR AgeToDays = 25568"
59090 Cnxn(0).Execute sql

59100 sql = "UPDATE BioTestDefinitions " & _
            "SET AgeToText = '80 Years', " & _
            "AgeToDays = 29200 WHERE AgeToDays = 29200 "
59110 Cnxn(0).Execute sql

59120 sql = "UPDATE BioTestDefinitions " & _
            "SET AgeToText = '120 Years', " & _
            "AgeToDays = 43830 WHERE AgeToDays >= 43800 "
59130 Cnxn(0).Execute sql

59140 Exit Sub

UpdateBioTestDefinitions_Error:

      Dim strES As String
      Dim intEL As Integer

59150 intEL = Erl
59160 strES = Err.Description
59170 LogError "modBioChemUpdate", "UpdateBioTestDefinitions", intEL, strES, sql


End Sub

