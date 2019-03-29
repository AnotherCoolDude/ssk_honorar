package main

import (
	"github.com/AnotherCoolDude/excel"
)

type monthlySummary struct {
	refSheet *excel.Sheet
}

func (ms *monthlySummary) Insert(sh *excel.Sheet) {
	currentCoords := ctx.customerSummary[revenue.string()][0]
	currentCoords.Column = 1
	currentCustomer := ms.refSheet.GetValue(currentCoords)
	cells := map[int]excel.Cell{
		1: excel.Cell{Value: currentCustomer, Style: excel.NoStyle()},
		2: excel.Cell{Value: ctx.customerSummary.formula(revenue).Reference(ms.refSheet.Name()).Add(), Style: excel.NoStyle()},
		3: excel.Cell{Value: ctx.customerSummary.formula(externalCosts).Reference(ms.refSheet.Name()).Add(), Style: excel.NoStyle()},
		4: excel.Cell{Value: ctx.customerSummary.formula(externalCostsChargeable).Reference(ms.refSheet.Name()).Add(), Style: excel.NoStyle()},
		5: excel.Cell{Value: ctx.customerSummary.formula(invoice).Reference(ms.refSheet.Name()).Add(), Style: excel.NoStyle()},
		6: excel.Cell{Value: ctx.customerSummary.formula(honorar).Reference(ms.refSheet.Name()).Add(), Style: excel.NoStyle()},
	}
	sh.AddRow(cells)
}


