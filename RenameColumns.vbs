Option Explicit

' Input and Output file paths
Dim inputFile, outputFile
inputFile = "C:\Users\ShamMahajan\Desktop\MF\ICRA\July Files\Debt_31072024.csv" 'Change this to your input CSV path
outputFile ="C:\Users\ShamMahajan\Desktop\MF\ICRA\July Files\Debt_31072024_1.csv"' Change this to your desired output CSV path

' Create objects for reading and writing files
Dim fso, inputFileObj, outputFileObj, textLine, columnNames
Set fso = CreateObject("Scripting.FileSystemObject")
Set inputFileObj = fso.OpenTextFile(inputFile, 1)
Set outputFileObj = fso.CreateTextFile(outputFile, True)

' Read the header line
textLine = inputFileObj.ReadLine
columnNames = Split(textLine, ",")

' Dictionary to track the count of each column name
Dim columnCountDict, colName, newHeader
Set columnCountDict = CreateObject("Scripting.Dictionary")

' Loop through each column name and rename if necessary
newHeader = ""
For Each colName In columnNames
    If columnCountDict.Exists(colName) Then
        columnCountDict(colName) = columnCountDict(colName) + 1
        newHeader = newHeader & colName & "_" & columnCountDict(colName) & ","
    Else
        columnCountDict.Add colName, 1
        newHeader = newHeader & colName & ","
    End If
Next

' Remove the trailing comma and write the new header to the output file
newHeader = Left(newHeader, Len(newHeader) - 1)
outputFileObj.WriteLine newHeader

' Write the remaining lines (data) to the output file
Do While Not inputFileObj.AtEndOfStream
    outputFileObj.WriteLine inputFileObj.ReadLine
Loop

' Close files
inputFileObj.Close
outputFileObj.Close

' Clean up
Set fso = Nothing
Set inputFileObj = Nothing
Set outputFileObj = Nothing
Set columnCountDict = Nothing
