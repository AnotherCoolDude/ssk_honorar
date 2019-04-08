package main

import (
	"github.com/AnotherCoolDude/excel"
)

func addProjectSummaryToSheet(sheet *excel.Sheet) {
	formulas := formulaFromIDs([]string{"revenue", "externalCosts", "externalCostsChargeable", "invoice"}, sheet)
	newRow := excel.Row{
		3: excel.NewCell(formulas[0].Add()),
		4: excel.NewCell(formulas[1].Add()),
		5: excel.NewCell(formulas[2].Add()),
		6: excel.NewCell(formulas[3].Add()),
	}
	sheet.AddRow(newRow)
}

func formulaFromIDs(ids []string, sheet *excel.Sheet) []excel.Formula {
	formulas := []excel.Formula{}
	for _, id := range ids {
		coords := []excel.Coordinates{}
		for _, cell := range sheet.Draft().CellsWithID(id) {
			coords = append(coords, cell.Coordinates())
		}
		formulas = append(formulas, excel.Formula{Coords: &coords})
	}
	return formulas
}

// type projectSummary struct {
// }

// func (pjts *projectSummary) Columns() []string {
// 	return []string{}
// }

// func (pjts *projectSummary) Insert(sh *excel.Sheet) {
// 	honorarFormula := excel.Formula{Coords: &[]excel.Coordinates{
// 		excel.Coordinates{Row: sh.NextRow(), Column: revenue.int()},
// 		excel.Coordinates{Row: sh.NextRow(), Column: invoice.int()},
// 		excel.Coordinates{Row: sh.NextRow(), Column: subsidiesEL.int()},
// 		excel.Coordinates{Row: sh.NextRow(), Column: subsidiesFK.int()},
// 	}}

// 	topBorderEuroStyle := excel.Style{Border: excel.Top, Format: excel.Euro}
// 	pjtsCells := map[int]excel.Cell{
// 		revenue.int():                 excel.Cell{Value: ctx.projectSummary.formula(revenue).Add(), Style: topBorderEuroStyle},
// 		externalCosts.int():           excel.Cell{Value: ctx.projectSummary.formula(externalCosts).Add(), Style: topBorderEuroStyle},
// 		externalCostsChargeable.int(): excel.Cell{Value: ctx.projectSummary.formula(externalCostsChargeable).Add(), Style: topBorderEuroStyle},
// 		invoice.int():                 excel.Cell{Value: ctx.projectSummary.formula(invoice).Sum(), Style: topBorderEuroStyle},
// 		honorar.int(): excel.Cell{Value: honorarFormula.Raw(func(coords []excel.Coordinates) string {
// 			return fmt.Sprintf("=%s-%s", coords[0].String(), coords[1].String())
// 		}), Style: topBorderEuroStyle},

// 		// set style for cells that are not filled out
// 		activity.int():   excel.Cell{Value: " ", Style: topBorderEuroStyle},
// 		fibu.int():       excel.Cell{Value: " ", Style: topBorderEuroStyle},
// 		paginiernr.int(): excel.Cell{Value: " ", Style: topBorderEuroStyle},
// 	}

// 	if len(sh.HeaderColumns()) > 10 {
// 		pjtsCells[subsidiesEL.int()] = excel.Cell{Value: ctx.projectSummary.formula(subsidiesEL).Sum(), Style: topBorderEuroStyle}
// 		pjtsCells[subsidiesFK.int()] = excel.Cell{Value: ctx.projectSummary.formula(subsidiesFK).Sum(), Style: topBorderEuroStyle}
// 		pjtsCells[honorar.int()] = excel.Cell{Value: honorarFormula.Raw(func(coords []excel.Coordinates) string {
// 			return fmt.Sprintf("=%s-%s-%s-%s", coords[0].String(), coords[1].String(), coords[2].String(), coords[3].String())
// 		}), Style: topBorderEuroStyle}
// 	}

// 	sh.AddRow(pjtsCells)
// 	ctx.customerSummary.addFromCurrentRow(sh, []header{
// 		revenue,
// 		externalCosts,
// 		externalCostsChargeable,
// 		invoice,
// 		honorar,
// 		subsidiesEL,
// 		subsidiesFK,
// 	})
// 	sh.AddEmptyRow()
// 	ctx.projectSummary = cellMap{}
// }
