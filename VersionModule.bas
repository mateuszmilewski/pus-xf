Attribute VB_Name = "VersionModule"
'The MIT License (MIT)
'
'Copyright (c) 2022 FORREST
' Mateusz Milewski mateusz@stellantis.com aka FORREST
'
'Permission is hereby granted, free of charge, to any person obtaining a copy
'of this software and associated documentation files (the "Software"), to deal
'in the Software without restriction, including without limitation the rights
'to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
'copies of the Software, and to permit persons to whom the Software is
'furnished to do so, subject to the following conditions:
'
'The above copyright notice and this permission notice shall be included in all
'copies or substantial portions of the Software.
'
'THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
'IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
'FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
'AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
'LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
'OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
'SOFTWARE.


'v006 - futher adjustments for input data -> using cost center as domain and magasin parameters
' final check on majNomFNR__xF -> PUS creation big mod
' v005 - some first new batch of info from PCV_SRC from DECORATOR - Sechel Sigapp  data from PCV_SRC 202212xx
' v004 - update on qty confirmed - pus is understanding this one
'   improtant note - grouping by month - PUS Master remembers confirmed qty but do not remember old data - if Sechel Sigapp will miss some data then
'   we will lose it! be careful!
'
' v003 - prog liv replaced with quasi PUS number
'   liste_pickupsheet_a_mettre_a_jour - temporary removed removed duplicates on column 6 is PUS creation - no prog liv (delivery program)
'   i need to think how I should work on it.
'
' v002 - there is a necessity for making change at the end
'    - even to perfrom any demo i need double check if the implementation already needs adjustment
'    - to catch to be def: TEMP_REMARK_FOR_DEMO
' v001 - starting point for xF version
