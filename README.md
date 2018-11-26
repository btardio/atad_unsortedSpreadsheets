# Using Key/Value Collections to Handle Unsorted Excel Sheets for Data Analytics

## Overview

For large worksheets in excel that are worked on by multiple people in organizations it could prove convenient to use a data structure to organize information for data analysis. In this tutorial we will go over the pros and cons of using the dictionary data structure and discuss alternatives.

## Pros

Dictionaries allow us to handle unsorted excel worksheets for instances where we have duplicate entries in a single column and are required to tally/aggregate those results. We could approach such a problem by iterating through the column data for all duplicate entries and performing data analysis but this would fail to provide accurate results unless the column was aggregated and sorted into groups based on duplicate values. Due to the row limit of excel spreadsheets this approach could be considered a feasible generic-computing solution for handling all Excel cases.

## Cons

For large datasets this requires iterating through all rows before iterating the aggregate groups, and even, depending on the dataset, iterating the entire dataset again, resulting in inefficiency. Excel does not allow us to make requests to APIs or SQL databases or access the row data in a way where we can specify the number of return results and view the results as pages. If physical memory space is a concern, on systems that do not have swap space, loading all values into a dictionary is impossible.

## Alternatives

The best alternative includes pre-sorting a worksheet. Other alternatives that could use less temporary storage space could include iterating the entire worksheet for each duplicate value in a column and freeing memory after each aggregate group is analyzed. A third alternative is probably unlikely to be used in worksheets but is worth mentioning and includes populating an SQLite database (which is stored as a file on the hard drive) and then querying the database from the VBA script.

## Process


### Dataset

For this tutorial we will be using the ICO (International Coffee Organization) Kaggle dataset, visit [this page](https://www.kaggle.com/sbajew/icos-crop-data).

### Data Analysis Requirement

We will impose a requirement that is to find the change in total production for each country from the record that is the oldest to the record that is the newest.

### An O(n^2) Solution

One solution could be to iterate the sorted dataset, sorted first by Country and then by Data, and find the 2000 date and once found, create another loop to find the 2010 figure. The code could look something like:


```vbnet

Sub total_change_O_n2()

    Dim lastCountry As String
    Dim visitedCountries As New Collection
    Dim foundcountry As Boolean
    Dim currentCountry As String
    Dim candidate As Date
    
    Dim minDate As Date
    Dim maxDate As Date
    
    Dim minDateRowIndex As Integer
    Dim maxDateRowIndex As Integer
    
    Dim pctChange As Double
    
    Dim divisor As Integer
    
    ' get number of rows
    nRows = Cells(Rows.Count, "A").End(xlUp).Row

    For i = 2 To nRows
        
        ' reset the foundcountry boolean
        foundcountry = False
        
        ' search through the visitedCountries list and check
        ' that we haven't already analyzed this country
        For Each itemi In visitedCountries
            If itemi = Cells(i, 1).Value Then
                foundcountry = True
                Exit For
            End If
        Next itemi
        
        ' if this country has not been analyzed, analyze
        If Not foundcountry Then
    
            ' set the value of the currentCountry we are analyzing
            currentCountry = Cells(i, 1).Value
            
            ' add the country to the list of analyzed countries
            visitedCountries.Add (Cells(i, 1).Value)
            
            ' set the first record as the min/max date
            minDate = (CDate((Cells(i, 3).Value) + " 1, " + _
                      ((Split(Cells(i, 2).Value, "/")(0)))))
            maxDate = (CDate((Cells(i, 3).Value) + " 1, " + _
                      ((Split(Cells(i, 2).Value, "/")(0)))))
            
            ' set the index of the record with the min/max date
            minDateRowIndex = i
            maxDateRowIndex = i
            
            ' iterate through all rows
            For j = 2 To nRows
            
                ' if the row is about the current country we are analyzing
                If (Cells(j, 1).Value = currentCountry) Then
                        
                    ' assign the date candidate which could be min or max date
                    candidate = (CDate((Cells(j, 3).Value) + " 1, " + _
                                ((Split(Cells(j, 2).Value, "/")(0)))))
                        
                    ' if this candidate is smaller than the current minimum date
                    If candidate < minDate Then
                        minDate = candidate ' set minDate to candidate
                        minDateRowIndex = j ' record the index of the row that
                                            ' has the minimum date found so far
                    End If
                    
                    ' if this candidate is larger than the maximum date
                    If candidate > maxDate Then
                        maxDate = candidate ' set maxDate to candidate
                        maxDateRowIndex = j ' record the index of the row that
                                            ' has the maximum date found so far
                    End If
                    
                End If
                
            Next j
            
            ' Incase the divisor is 0 set it to 1
            divisor = Cells(minDateRowIndex, 4)
            If divisor = 0 Then divisor = 1
            
            ' calculate percent change
            pctChange = CDbl((Cells(maxDateRowIndex, 4) - _
                        Cells(minDateRowIndex, 4)) / divisor) * 100
            
            ' Debug print the pctChange
            Debug.Print ("Country: " + currentCountry)
            Debug.Print ("Percent Change: %" + _
                         Format(CStr(pctChange), "##,##0.00"))
            Debug.Print ("###")
    
        End If
    Next i
End Sub

```

Because we have introduced a nested loop we roughly analyze this algorithm to be O(n^2). Theoretically, this algorithm could read through the entire dataset to find the 2000 figure (which may be the last row) and then read through the remaining dataset, which could be the entire dataset - 1, to find the 2010 figure. This solution is considered O(n^2) worst-case. There are numerous explanations of big O notation, applied to both comp/sci and mathematics. I have listed a few below for further reading.

* A must-read overview, [this page](http://www-bcf.usc.edu/~stejada/csci101/slides/2012/SortingAlgoritms.pdf).
* The famous Tower of Hanoi, [this page](http://www-bcf.usc.edu/~stejada/csci101/slides/2012/TowerHanoi.pdf).
* The Wikipedia entry for Analysis of Algorithms, [this page](https://en.wikipedia.org/wiki/Analysis_of_algorithms).
* A short test-yourself resource, [this page](http://pages.cs.wisc.edu/~vernon/cs367/notes/3.COMPLEXITY.html).



### A better O(n) Solution


This solution builds a key / value collection which is indexable by key. This indexing is an O(n) operation. This solutions keeps another collection of all countries, which is less than or equal to the size of all rows in the spreadsheet (in this dataset it is much less than the number of rows). The solution iterates the spreadsheet only once and then iterates the countries only once. The rough worst-case scenario analysis of this algorithm is O(2n) or simply, O(n) if you wish to drop constants.

This solution uses a class module that looks like. Alternatively we could have built a third (min) and fourth (max) collection with country as key and row number as value.

```vbnet
' rowdateclass Class Module

Private adate As Date
Private row As Integer

Public Function setDate(inDate As Date)
    adate = inDate
End Function

Public Function setRow(inRow As Integer)
    row = inRow 
End Function

Public Function getDate() As Date
    getDate = adate 
End Function
 
Public Function getRow() As Integer 
    getRow = row 
End Function

```

The sub procedure looks like:

```vbnet

Sub total_change_O_n()

    Debug.Print ("O(n)")

    Dim minDate As New Collection
    Dim maxDate As New Collection
    
    Dim countries As New Collection
    Dim foundCountry As Boolean
    Dim divisor As Integer
    Dim rowdateinstance As New rowdateclass
        
    Dim aMinDateRowIndex As Integer
    Dim aMaxDateRowIndex As Integer
    
    ' get number of rows
    nRows = Cells(Rows.Count, "A").End(xlUp).row

    ' iterate all rows of the spreadsheet, only once
    For i = 2 To nRows
        
        ' reset the foundcountry boolean
        foundCountry = False
        
        ' set the country for the current row
        countryofrow = Cells(i, 1).Value
        
        ' search through the visitedCountries list and check
        ' that we haven't added this country
        For Each itemi In countries
            If itemi = countryofrow Then
                foundCountry = True
                Exit For
            End If
        Next itemi
        
        ' re-initialize a new rowdateclass for the row we are on
        Set rowdateclass = New rowdateclass
        
        ' assign the rowdateclass values
        rowdateclass.setDate (CDate((Cells(i, 3).Value) + " 1, " + _
                             ((Split(Cells(i, 2).Value, "/")(0)))))
        rowdateclass.setRow (i)
        
        If Not foundCountry Then
            ' add the country to the list of all countries
            countries.Add (countryofrow)
            minDate.Add rowdateclass, countryofrow
            maxDate.Add rowdateclass, countryofrow

        End If
        
        If foundCountry Then
            ' check to see if the min max date is less or greater
            ' if it is then remove the rowdateclass instance from the
            ' collection and add the new one
            If minDate.Item(countryofrow).getDate > rowdateclass.getDate() Then
                minDate.Remove (countryofrow)
                minDate.Add rowdateclass, countryofrow
            Else
                If maxDate.Item(countryofrow).getDate < rowdateclass.getDate() Then
                    maxDate.Remove (countryofrow)
                    maxDate.Add rowdateclass, countryofrow
                End If
            End If
            
        End If
        
    Next i
    
    ' now that we have a collection that is indexable by country we can
    ' analyze our results as we did in the o(n^2) solution
    For Each country In countries
        
        aMinDateRowIndex = minDate(country).getRow()
        aMaxDateRowIndex = maxDate(country).getRow()
        
        ' Incase the divisor is 0 set it to 1
        divisor = Cells(aMinDateRowIndex, 4)
        If divisor = 0 Then divisor = 1
        
        ' calculate percent change
        pctChange = CDbl((Cells(aMaxDateRowIndex, 4) - _
                    Cells(aMinDateRowIndex, 4)) / divisor) * 100
        
        ' Debug print the pctChange
        Debug.Print ("Country: " + country)
        Debug.Print ("Percent Change: %" + _
                     Format(CStr(pctChange), "##,##0.00"))
        Debug.Print ("###")
        
    Next country
    
End Sub

```

### Conclusion

I hope you enjoyed reading through this tutorial. I also hope you have the chance to take a look at a few of the big-O notation resources.
