# libreoffice-to-pdf
A tool for converting Calc files in PDF files
Contact me walfrido_15@hotmail.com

Utilizado para versoes > 4 onde nao tem a pasta URE
SetEnvironmentVariable("UNO_PATH", unoPath, EnvironmentVariableTarget.Process)
SetEnvironmentVariable("PATH", GetEnvironmentVariable("PATH") + ";" + unoPath, EnvironmentVariableTarget.Process)

Utilizado para versos < 4

Versões que não existe a pasta 'URE' o PATH ficará na pasta 'program'
Tendo a pasta URE o PATH será a pasta 'URE'

Dim unoPath = Configuration.ConfigurationManager.AppSettings("UNO_PATH").ToString() 

C:\Program Files (x86)\LibreOffice 3.6\program

Dim urePath = Configuration.ConfigurationManager.AppSettings("URE_PATH").ToString()

C:\Program Files (x86)\LibreOffice 3.6\URE\bin
