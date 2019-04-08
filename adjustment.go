package main

import ()

type adjustment struct {
	jobnr    string
	year     string
	amountEL float32
	amountFK float32
}

// func (adj *adjustment) Insert(sh *excel.Sheet) {
// 	adjCells := map[int]excel.Cell{
// 		job.int():      excel.Cell{Value: adj.jobnr, Style: excel.NoStyle()},
// 		year.int():     excel.Cell{Value: adj.year, Style: excel.NoStyle()},
// 		amountEL.int(): excel.Cell{Value: adj.amountEL, Style: excel.EuroStyle()},
// 		amountFK.int(): excel.Cell{Value: adj.amountFK, Style: excel.EuroStyle()},
// 	}
// 	sh.AddRow(adjCells)
// 	ctx.adjustmentSummary.addFromCurrentRow(sh, []header{amountEL, amountFK})
// }

// type adjustmentSummary struct {
// }

// func (adjs *adjustmentSummary) Insert(sh *excel.Sheet) {
// 	topBorderNoStyle := excel.Style{Border: excel.Top, Format: excel.NoFormat}
// 	topBorderEuroStyle := excel.Style{Border: excel.Top, Format: excel.Euro}
// 	smyCells := map[int]excel.Cell{
// 		job.int():      excel.Cell{Value: "Gesamt", Style: topBorderNoStyle},
// 		year.int():     excel.Cell{Value: " ", Style: topBorderNoStyle},
// 		amountEL.int(): excel.Cell{Value: ctx.adjustmentSummary.formula(amountEL).Add(), Style: topBorderEuroStyle},
// 		amountFK.int(): excel.Cell{Value: ctx.adjustmentSummary.formula(amountFK).Add(), Style: topBorderEuroStyle},
// 	}
// 	sh.AddRow(smyCells)
// }
