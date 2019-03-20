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
		excel.Coordinates{Row: sh.NextRow(), Column: revenue.int() + 1},
		excel.Coordinates{Row: sh.NextRow(), Column: invoice.int() + 1},
		excel.Coordinates{Row: sh.NextRow(), Column: subsidiesEL.int() + 1},
		excel.Coordinates{Row: sh.NextRow(), Column: subsidiesFK.int() + 1},
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

	if _, ok := ctx.projectSummary[subsidiesEL.string()]; ok {
		pjtsCells[subsidiesEL.int()] = excel.Cell{Value: ctx.projectSummary.formula(subsidiesEL).Sum(), Style: topBorderEuroStyle}
		pjtsCells[subsidiesFK.int()] = excel.Cell{Value: ctx.projectSummary.formula(subsidiesFK).Sum(), Style: topBorderEuroStyle}
		pjtsCells[honorar.int()] = excel.Cell{Value: honorarFormula.Raw(func(coords []excel.Coordinates) string {
			return fmt.Sprintf("=%s-%s-%s-%s", coords[0].ToString(), coords[1].ToString(), coords[2].ToString(), coords[3].ToString())
		}), Style: topBorderEuroStyle}
	}

	sh.AddRow(pjtsCells)
	ctx.customerSummary.addFromCurrentRow(sh, []header{
		revenue,
		externalCosts,
		externalCostsChargeable,
		invoice,
		honorar,
		subsidiesEL,
		subsidiesFK,
	})
	sh.AddEmptyRow()
	ctx.projectSummary = cellMap{}
}
