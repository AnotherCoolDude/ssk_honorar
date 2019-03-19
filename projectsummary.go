package main

import (
	"fmt"
	"github.com/AnotherCoolDude/excel"
)

type projectSummary struct {
}

func (pjts *projectSummary) Columns() []string {
	return []string{}
}

func (pjts *projectSummary) Insert(sh *excel.Sheet) {
	honorarFormula := excel.Formula{Coords: []excel.Coordinates{
		excel.Coordinates{Row: sh.NextRow(), Column: revenue.int()},
		excel.Coordinates{Row: sh.NextRow(), Column: invoice.int()},
	}}

	topBorderEuroStyle := excel.Style{Border: excel.Top, Format: excel.Euro}
	pjtsCells := map[int]excel.Cell{
		revenue.int():                 excel.Cell{Value: ctx.projectSummary.formula(revenue).Add(), Style: topBorderEuroStyle},
		externalCosts.int():           excel.Cell{Value: ctx.projectSummary.formula(externalCosts).Add(), Style: topBorderEuroStyle},
		externalCostsChargeable.int(): excel.Cell{Value: ctx.projectSummary.formula(externalCostsChargeable).Add(), Style: topBorderEuroStyle},
		invoice.int():                 excel.Cell{Value: ctx.projectSummary.formula(invoice).Sum(), Style: topBorderEuroStyle},
		honorar.int(): excel.Cell{Value: honorarFormula.Raw(func(coords []excel.Coordinates) string {
			return fmt.Sprintf("=%s-%s", coords[0].ToString(), coords[1].ToString())
		}), Style: topBorderEuroStyle},

		// set style for cells that are not filled out
		activity.int():   excel.Cell{Value: " ", Style: topBorderEuroStyle},
		fibu.int():       excel.Cell{Value: " ", Style: topBorderEuroStyle},
		paginiernr.int(): excel.Cell{Value: " ", Style: topBorderEuroStyle},
	}
	sh.AddRow(pjtsCells)
	ctx.customerSummary.addFromCurrentRow(sh, []header{
		revenue,
		externalCosts,
		externalCostsChargeable,
		invoice,
		honorar,
	})
	sh.AddEmptyRow()
	ctx.projectSummary = cellMap{}
}
