Import-Module ImportExcel
$expenses = Import-Excel -Path 'C:\temp\wings-expenses.xlsx' -WorksheetName Expenses -StartRow 4
$expenses.Category | Sort-Object -Unique | ForEach-Object{
	# Category Breakdown
	$thisCategory = $_
	$theseEntries = $expenses | Where-Object{ $_.Category -eq $thisCategory }
	$categoryAmount = $theseEntries.'Actual Extended' | Measure-Object -Sum
	('{0} [{1}] {2}' -f 'Actual', $categoryAmount.sum, $thisCategory )

	# Subcategory Breakdown
	($theseEntries).Subcategory | Sort-Object -Unique | ForEach-Object{
		$thisSubcategory = $_
		$theseSubEntries = $theseEntries | Where-Object{ $_.Subcategory -eq $thisSubcategory }
		$subcategoryAmount = $theseSubEntries.'Actual Extended' | Measure-Object -Sum
		('{0} [{1}] {2}' -f $thisCategory, $subcategoryAmount.sum, $thisSubcategory )

		# Related Breakdown
		($theseSubEntries).Related | Sort-Object -Unique | ForEach-Object{
			$thisRelation = $_
			$theseRelatedEntries = $theseSubEntries | Where-Object{ $_.Related -eq $thisRelation }
			$relatedAmount = $theseRelatedEntries.'Actual Extended' | Measure-Object -Sum
			('{0} [{1}] {2}' -f $thisSubcategory, $relatedAmount.sum, $thisRelation )
		}
	}
} | Clip