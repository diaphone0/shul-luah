# vbzmanim
VBA &amp; VB6 modules for zmanim &amp; hebrew calendar, ported from [yparitcher/libzmanim](https://github.com/yparitcher/libzmanim),
which is a port of [KosherJava/Zmanim](https://github.com/KosherJava/zmanim) developed by Eliyahu Hershfeld. [https://kosherjava.com](https://kosherjava.com)

The Daf-Yomi related code was ported from https://github.com/NykUser/MyZman/

## Usage:
Include the modules in your Visual Basic project ("mod_*" files) for using the calendar and zmanim functions.

See below a brief summary on using the calendar and zmanim functions.

## The hdate data type:
In Visual Basic, working with gregorian dates involves the builtin "Date" data type.

For working with hebrew dates a dedicated type `hdate` is being used, which has simillar elements to "struct tm" type in C (year, month etc...) with some dedicated elements for the hebrew calendar (leap, EY and offset).

This type can be directly initialized or to be converted from VB's standard gregorian `Date` type, using the `ConvertDate` function.
A hdate variable can be converted back into standard `Date` type using the `HDateGregorian` function.

### Example 1 - Converting gregorian to hebrew dates and vice versa
```VBA
Dim gre_d As Date
Dim heb_d As hdate
Dim gre_d2 As Date

gre_d = Date
MsgBox "The current gregorian date - " & gre_d
heb_d = ConvertDate(gre_d)
MsgBox "converted to hebrew date - " & HDateFormat(heb_d) 'HDateFormat converts a hebrew date to string
gre_d2 = HDateGregorian(heb_d)
MsgBox "converted back to gregorian date - " & gre_d2
```
### Getting additional calendar info for a hdate
For getting additional info for a specific hdate, see the additional 'Get...' and 'Is...' functions listed in `mod_hebrewcalendar`
The most relevant functions are:
* GetParshah - returns the parshah for provided shabbos in `parshah` enum type
* GetYomTov - returns the moed or yomtov for provided hdate in `yomtov` enum type
* GetMolad - returns the molad details for the provided month and year, as `hdate` data type
* GetOmer - if relevant
* IsCandleLighting - true if provided hdate is erev shabbos / yomtov etc
* IsAssurBeMelachah - true if provided hdate is shabbos / yomtov etc

See descriptions for each function on their headers in the code, and also see usage examples in the example below and the included samples!
### Formatting hebrew dates and related info to strings
The hebrew dates are represented as numeric values under `hdate` data type and the related enums `parshah` and `yomtov`.
This data can be formatted as string using the functions listed in `mod_hdateformat` The most relevant functions are:
* HDateFormat / HDateOrFormat - convert hebrew date to string
* YomTovFormat - get the title of the sepcified `yomtov` enum value for the relevant YomTov / Moed (including special shabbasos)
* ParshahFormat - get the title of the parshah in the specified `parshah` enum value
* MoladFormat - formats to string the specified molad info specified in a hdate (that is retrieved using GetMolad) 
* NumToHChar - converts a number to hebrew char representation

See descriptions for each function on their headers in the code, and also see usage examples in the example below and the included samples!

### Example 2 - Getting the next Parshah or YomTov
```VBA
Dim heb_date As hdate
Dim parsh As parshah, ytov As yomtov
Dim shabbos_title As String
heb_date = ConvertDate(Date)
'find next shabbos day, increment days until reaching a day which is AssurBeMelachah
Do While IsAssurBeMelachah(heb_date) = 0
    Call HDateAddDay(heb_date, 1)
Loop

'get the parsha (as enum value)
parsh = GetParshah(heb_date)
If parsh <> NOPARSHAH Then
    'if the found date has a parshah then add it's parsha to the string
    'note - ParshahFormat converts parshah enum values to their titles in string
    shabbos_title = "שבת פרשת " & ParshahFormat(parsh)
Else
    'if the found date has no parsha then it is probably a yomtov or shabbos chol ha'moed
    ytov = GetYomTov(heb_date)
    'convert the found date from a yomyov enum value type to string using YomTovFormat
    If ytov <> CHOL Then shabbos_title = YomTovFormat(ytov)
End If

'check if the found date is a spacial shabbos (hagadol, 4 parshios etc)
If GetSpecialShabbos(heb_date) <> CHOL Then
    shabbos_title = shabbos_title & vbCrLf & YomTovFormat(GetSpecialShabbos(heb_date))
End If
'show the result
MsgBox shabbos_title
```

## Zmanim functions:
All the relevant functions for calculating zmanim are listed in `mod_zmanim`. See examples in the included samples

## Limud Yomi functions:
The relevant functions for limud yomi are listed in `mod_shiur` and `mod_dafyomi`. See examples in the included samples
