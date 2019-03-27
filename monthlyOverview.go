package main

import (
	"fmt"
	"github.com/AnotherCoolDude/excel"
)

type monthlyOverview struct {
	refSheet *excel.Sheet
}

func (ms *monthlyOverview) Insert(sh *excel.Sheet) {
	abbrevation := ctx.monthlyOverview.formula(customer).Raw(func(coords []excel.Coordinates) string {
		newCoords := coords[0]
		newCoords.Column = 2
		newCoords.Row -= 3
		jobnr := fmt.Sprintf("%s", ms.refSheet.GetValue(newCoords))
		return jobnr[:4]
	})
	cells := map[int]excel.Cell{
		1: excel.Cell{Value: ctx.monthlyOverview.formula(customer).Reference(ms.refSheet.Name()).Add(), Style: excel.NoStyle()},
		2: excel.Cell{Value: abbrevation, Style: excel.NoStyle()},
		3: excel.Cell{Value: ctx.monthlyOverview.formula(revenue).Reference(ms.refSheet.Name()).Add(), Style: excel.EuroStyle()},
		4: excel.Cell{Value: ctx.monthlyOverview.formula(externalCosts).Reference(ms.refSheet.Name()).Add(), Style: excel.EuroStyle()},
		5: excel.Cell{Value: ctx.monthlyOverview.formula(externalCostsChargeable).Reference(ms.refSheet.Name()).Add(), Style: excel.EuroStyle()},
		6: excel.Cell{Value: ctx.monthlyOverview.formula(invoice).Reference(ms.refSheet.Name()).Add(), Style: excel.EuroStyle()},
		7: excel.Cell{Value: ctx.monthlyOverview.formula(honorar).Reference(ms.refSheet.Name()).Add(), Style: excel.EuroStyle()},
	}
	sh.AddRow(cells)
	ctx.monthlyOverview = cellMap{}
}
