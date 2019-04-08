package main

import (
	"github.com/AnotherCoolDude/excel"
)

func extract() (rentData, erData, abgr18Data, adj18Data [][]string) {
	rentExcel := excel.File(rentabilität, "", false)
	erExcel := excel.File(eingangsrechnungen, "", false)
	abgr18Excel := excel.File(abgr19, "", false)
	adj18Excel := excel.File(adj19, "", false)

	rentData = rentExcel.FirstSheet().ExtractColumns([]string{
		"A", "C", "E", "G", "I", "L", "E",
	})

	erData = erExcel.FirstSheet().ExtractColumns([]string{
		"A", "F", "G", "I", "K",
	})

	abgr18Data = abgr18Excel.Sheet("konsolidiert").ExtractColumns([]string{
		"A",
	})

	adj18Data = adj18Excel.Sheet("konsolidiert18").ExtractColumns([]string{
		"A", "B", "C",
	})
	return
}
