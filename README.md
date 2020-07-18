# vba_excel-Sort_Listbox_by_Columns
Module in vba to Order a LISTBOX with any number of columns using an Arrays based strategy

The module can be imported to the Excel Proyect.
Internally it has three functions:
  1) SortListBox     --> by J.G.Arvidsson (2020-07-18).
	
  2) SortArrayAtoZ   --> From page ExcelOffTheGrid.
	
  3) SortArrayZtoA   --> From page ExcelOffTheGrid.

* About points 2 & 3: I do not know the authorship, if you have any additional information, comment on it. Thank you.

Once installed the module, only need the following order:

	SortListBox(ListBoxName As MSForms.ListBox, Ascendent As Boolean, Optional SortByColumn As Double = 0)

Where the parameters indicated:
- LisBoxName = name of the ListBox that want be sort.
- Ascendent = True (from A to Z) or False (from Z to A).
- SortByColumn = Column number where you want to sort by. This value is optional. If don't write nothing the Listbox will be ordered by first column, if the value is highest that the column number existent, it value will be the last column existent.
