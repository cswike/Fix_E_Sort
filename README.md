# Fix_E_Sort
Allows part numbers with the format [number]E[number] to be sorted in Excel as alphanumeric strings, rather than numbers in scientific/exponential notation.

------------------------------------------------

Excel reads part numbers like "9E12" as a numeric value using scientific notation (in this case, it reads 9E12 as 9000000000000, or 9 with 12 zeroes).

This is a problem when sorting a list, as strings like "9E12" will be sorted as numbers rather than alphanumerically - e.g. a sorted list might look like:

{1, 900, 9E12, 9A12, 9B12, 9C12, 9D12, 9F12}

where we really want the 9E12 to come after 9D12.

While you can tell Excel to treat your part number as text by either A.) prepending a single apostrophe, i.e. '9E12 or B.) formatting the cell/range containing your alphanumeric strings to Text, I found that this does not apply to sorting the data - Excel will still sort 9E12 together with the other numbers, even when you tell Excel to treat it as text.

------------------------------------------------

This macro adds a helper column for sorting, and appends a single parenthesis ( after the E for part numbers with this format.
Since ( is a special character, it gets sorted before any numbers or letters, and this also forces Excel to read it as a string instead of a number in E-notation.
Also, ( doesn't have any mathematical operations assigned to it, so it should be safe to use.

Once the macro finishes, just sort by the helper column and your list should be sorted alphanumerically like you wanted.
You can delete the helper column when your list is sorted the way you want.
