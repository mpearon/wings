# https://sankeymatic.com/build/
if($IsWindows){
	$path = '~\Documents\temp\autoscrub\wings-expenses.xlsx'
}
elseIf($IsLinux){
	$path = '~/Documents/temp/autoscrub/wings-expenses.xlsx'
}
elseIf($IsMacOS){
	$path = '~/Documents/temp/autoscrub/wings-expenses.xlsx'
}
Import-Module ImportExcel
$expenses = Import-Excel -Path $path -WorksheetName Expenses -StartRow 4
$expenses.Category | Sort-Object -Unique | ForEach-Object{
	# Category Breakdown
	$thisCategory = $_
	$theseEntries = $expenses | Where-Object{ $_.Category -eq $thisCategory }
	$categoryAmount = $theseEntries.'Actual Extended' | Measure-Object -Sum
	('{0} [{1}] {2}' -f 'Actual', [Math]::Round($categoryAmount.sum,2), $thisCategory )

	# Subcategory Breakdown
	($theseEntries).Subcategory | Sort-Object -Unique | ForEach-Object{
		$thisSubcategory = $_
		$theseSubEntries = $theseEntries | Where-Object{ $_.Subcategory -eq $thisSubcategory }
		$subcategoryAmount = $theseSubEntries.'Actual Extended' | Measure-Object -Sum
		('{0} [{1}] {2}' -f $thisCategory, [Math]::Round($subcategoryAmount.sum,2), $thisSubcategory )

		# Related Breakdown
		($theseSubEntries).Related | Sort-Object -Unique | ForEach-Object{
			$thisRelation = $_
			$theseRelatedEntries = $theseSubEntries | Where-Object{ $_.Related -eq $thisRelation }
			$relatedAmount = $theseRelatedEntries.'Actual Extended' | Measure-Object -Sum
			('{0} [{1}] {2}' -f $thisSubcategory, [Math]::Round($relatedAmount.sum,2), $thisRelation )
		}
	}
}