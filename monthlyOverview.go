package main

import (
	"fmt"
	"github.com/AnotherCoolDude/excel"
)

type monthlyOverview struct {
	refSheet *excel.Sheet
}

func (ms *monthlyOverview) Insert(sh *excel.Sheet) {
	fmt.Printf("currentRow: %d\n", ms.refSheet.CurrentRow())
	fmt.Printf("Value: %s\n", ms.refSheet.GetValue(excel.Coordinates{Row: ms.refSheet.CurrentRow(), Column: 1}))
	cells := map[int]excel.Cell{
		1: excel.Cell{Value: ms.refSheet.GetValue(excel.Coordinates{Row: ms.refSheet.CurrentRow(), Column: 1}), Style: excel.NoStyle()},
		2: excel.Cell{Value: ctx.monthlyOverview.formula(revenue).Reference(ms.refSheet.Name()).Add(), Style: excel.NoStyle()},
		3: excel.Cell{Value: ctx.monthlyOverview.formula(externalCosts).Reference(ms.refSheet.Name()).Add(), Style: excel.NoStyle()},
		4: excel.Cell{Value: ctx.monthlyOverview.formula(externalCostsChargeable).Reference(ms.refSheet.Name()).Add(), Style: excel.NoStyle()},
		5: excel.Cell{Value: ctx.monthlyOverview.formula(invoice).Reference(ms.refSheet.Name()).Add(), Style: excel.NoStyle()},
		6: excel.Cell{Value: ctx.monthlyOverview.formula(honorar).Reference(ms.refSheet.Name()).Add(), Style: excel.NoStyle()},
	}
	sh.AddRow(cells)
	ctx.monthlyOverview = cellMap{}
}
