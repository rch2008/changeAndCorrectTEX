md %1
copy %1.docx.tmp\word\media\*.*  %1
rd /s /Q %1.debug
rd /s /Q %1.docx.tmp
del /s /Q %1.csv
del /s /Q %1.log
del /s /Q %1.xml