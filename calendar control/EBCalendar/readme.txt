EBCalendar User Control

Introduction:
This User Control is designed to fit the needs of the coder who needs a date entry field without the overhead of having to ship (and install) the Microsoft Windows Common Controls-2 OCX.

The control is a self contained User Control which can be added to any project or compiled as a stand alone OCX.

Properties:
The control has only two properties:

Text (Date)
	The currently selected Date

AutoSelect (Boolean)
	When set, automatically selects the date when the control receives the focus
	
Notes:
The control has two minor issues which shouldn't affect its usefulness:

Ensure their is enough room on the form under the control to display the drop down calendar otherwise the calendar will be clipped.

If the control is placed above other controls on the form, ensure you set it's Z-Order to 0 (topmost) otherwise the drop-down calendar will appear behind your other controls on the form.  You can set the Z-Order by selecting the control on your form, right-clicking and select "Bring To Front".

Contact:
If you have any comments, suggestions or improvements please send them to me at:
RichardAllsebrook@earlybirdmarketing.com

You will, of course, receive full credit for your contribution.