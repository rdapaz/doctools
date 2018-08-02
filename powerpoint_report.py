import os
import win32com.client
import sys

my_dir = r'C:\Users\rdapaz\Desktop\Temp\BTS PSRs'

for subdir, dirs, files in os.walk(my_dir):
    for file in files:
        print(file, subdir, sep="|")


# Public Sub Test()
# Dim obj As PowerPoint.Shape

# Dim pjt As PowerPoint.Presentation

# Set pjt = PowerPoint.Application.ActivePresentation


# For Each shp In pjt.Slides(1).Shapes
#     If Left(shp.Name, 5) = "Table 61" Then
#         Debug.Print shp.Name
#         Debug.Print shp.Table.Cell(1, 1).Shape.TextFrame.TextRange.Text
#         Debug.Print
#         Debug.Print
#     End If
# Next shp


C:\Users\rdapaz\AppData\Local\Temp\gen_py

# End Sub
