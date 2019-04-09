package main

import (
	"github.com/AnotherCoolDude/excel"
)

type adjustedProject struct {
	customer                string
	jobnr                   string
	revenue                 float32
	externalCosts           float32
	externalCostsChargeable float32
	invoice                 []float32
	activity                []string
	fibu                    []string
	paginiernr              []string
	subsidiesYear           []string
	subsidiesEL             []float32
	subsidiesFK             []float32
	honorar                 float32
}

func (adj *adjustedProject) addRowToSheet(sheet *excel.Sheet) {
	newRow := excel.Row{
		1: excel.NewCell("").SetID(adj.customer),
		2: excel.NewCell(adj.jobnr).SetID("jobnr"),
		3: excel.NewEuroCell(adj.revenue).SetID("revenue"),
		4: excel.NewEuroCell(adj.externalCosts).SetID("externalCosts"),
		5: excel.NewEuroCell(adj.externalCostsChargeable).SetID("externalCostsChargeable"),
	}
	sheet.AddRow(newRow)

	for i := range adj.invoice {
		newRow = excel.Row{
			6: excel.NewEuroCell(adj.invoice[i]).SetID("invoice"),
			7: excel.NewCell(adj.activity[i]),
			8: excel.NewCell(adj.fibu[i]),
			9: excel.NewCell(adj.paginiernr[i]),
		}
		sheet.AddRow(newRow)
	}

	for i := range adj.subsidiesEL {
		newRow = excel.Row{
			11: excel.NewEuroCell(adj.subsidiesEL[i]).SetID("subsidiesEL"),
			12: excel.NewEuroCell(adj.subsidiesFK[i]).SetID("subsidiesFK"),
			13: excel.NewCell(adj.subsidiesYear[i]).SetID("subsidiesYear"),
		}
		sheet.AddRow(newRow)
	}

}

// func (adjpjt *adjustedProject) Insert(sh *excel.Sheet) {
// 	adjpjtCells := map[int]excel.Cell{
// 		jobnr.int():                   excel.Cell{Value: adjpjt.jobnr, Style: excel.NoStyle()},
// 		revenue.int():                 excel.Cell{Value: adjpjt.revenue, Style: excel.EuroStyle()},
// 		externalCosts.int():           excel.Cell{Value: adjpjt.externalCosts, Style: excel.EuroStyle()},
// 		externalCostsChargeable.int(): excel.Cell{Value: adjpjt.externalCostsChargeable, Style: excel.EuroStyle()},
// 	}
// 	sh.AddRow(adjpjtCells)
// 	ctx.projectSummary.addFromCurrentRow(sh, []header{jobnr, revenue, externalCosts, externalCostsChargeable})

// 	cells := map[int]excel.Cell{}
// 	for i, inv := range adjpjt.invoice {
// 		cells = map[int]excel.Cell{
// 			invoice.int():    excel.Cell{Value: inv, Style: excel.EuroStyle()},
// 			activity.int():   excel.Cell{Value: adjpjt.activity[i], Style: excel.NoStyle()},
// 			fibu.int():       excel.Cell{Value: adjpjt.fibu[i], Style: excel.NoStyle()},
// 			paginiernr.int(): excel.Cell{Value: adjpjt.paginiernr[i], Style: excel.NoStyle()},
// 		}
// 		sh.AddRow(cells)
// 		ctx.projectSummary.addFromCurrentRow(sh, []header{invoice})
// 	}
// 	// get lenght of the longer subsidies slice
// 	lenghtSubsidies := len(adjpjt.subsidiesFK)
// 	if len(adjpjt.subsidiesEL) > len(adjpjt.subsidiesFK) {
// 		lenghtSubsidies = len(adjpjt.subsidiesEL)
// 	}

// 	for i := 0; i < lenghtSubsidies; i++ {
// 		cells = map[int]excel.Cell{
// 			13:                excel.Cell{Value: adjpjt.subsidiesYear[i], Style: excel.NoStyle()},
// 			subsidiesEL.int(): excel.Cell{Value: adjpjt.subsidiesEL[i], Style: excel.EuroStyle()},
// 			subsidiesFK.int(): excel.Cell{Value: adjpjt.subsidiesFK[i], Style: excel.EuroStyle()},
// 		}
// 		sh.AddRow(cells)
// 		ctx.projectSummary.addFromCurrentRow(sh, []header{subsidiesEL, subsidiesFK})
// 	}

//}
