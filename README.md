# CAN DBC to Excel converter v3.6

Note:
* Busload related columns [A-F] are added in message sheet.
* In Busload calculation formula, busload value is taken as 1.05
* If Busload value is changed, then change 1.05 to new value and reapply the formula.

Updates done:

v3.3:
* Busload related columns [A-F] are added in message sheet

v3.4:
* Filter will be applied to all columns automatically
* '-' will be added to blank cells in Unit and Invalid value columns in Signals sheet
* Updated error messages

v3.5:
* In “messages” sheet Cell ‘R1’ value changed from “Message ID [dez]” to “Message ID [dec]”
* In “signals” sheet the Cell ‘E1’ value changed from “Message ID [Dec]” to “Message ID [hex]”
* In “signals” sheet the Cell ‘V1’ value changed from “Multiplex Value [dez]” to “Multiplex Value [dec]”
* Removed the marker from “signals” and “message” sheets
* “Number Stored as Text issue” solved

v3.6:
* Table format has been changed.
