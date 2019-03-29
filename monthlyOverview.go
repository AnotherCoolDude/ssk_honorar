package main

import (
	"fmt"
	"github.com/AnotherCoolDude/excel"
)

type monthlyOverview struct {
	refSheet *excel.Sheet
}

func (ms *monthlyOverview) Insert(sh *excel.Sheet) {
	abbrevation := ctx.monthlyOverview.formula(revenue).Raw(func(coords []excel.Coordinates) string {
		fmt.Println(coords)
		newCoords := coords[0]
		newCoords.Column--
		jobnr := fmt.Sprintf("%s", ms.refSheet.GetValue(newCoords))
		fmt.Println(jobnr)
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

func (ms *monthlyOverview) in(sh *Sheet) {
	overviewRows := [][]excel.Cell
	abbr := ""
	for _, row := range ms.refSheet.Draft() {
		if row[1].Value != excel.DraftCell {
			abbr = fmt.Sprintf("%s", row[1].Value)[:4]
		}
		if row[0].Value != excel.DraftCell {
			sh.AddRow(map[int][]excel.Cell{
				1: row[0].ChangeStyle(excel.NoStyle()),
				2: excel.Cell{Value: abbr, Style: excel.NoStyle()},
			})
		}
	}
}

func (ms *monthlyOverview) insert(sh *excel.Sheet) {
	refDraft := ms.refSheet.Draft()[1:]
	summaryRows := [][]excel.Cell{}
	currentAbbr := ""
	for i, row := range refDraft {
		if row[1].Value != excel.DraftCell {
			currentAbbr = fmt.Sprintf("%s", row[1].Value)[:4]
		}
		if row[0].Value != excel.DraftCell {
			row[1].Value = currentAbbr
			summaryRows = append(summaryRows, row)
		}
	}

	for _, row := range summaryRows {
		sh.AddRow(map[int]excel.Cell{
			1: excel.Cell{Value: row[0].StringWithReference(sh), Style: excel.NoStyle()},
			2: excel.Cell{Value: row[1].StringWithReference(sh), Style: excel.NoStyle()},
			3: excel.Cell{Value: ctx.monthlyOverview.formula(revenue).Reference(ms.refSheet.Name()).Add(), Style: excel.EuroStyle()},
			4: excel.Cell{Value: ctx.monthlyOverview.formula(externalCosts).Reference(ms.refSheet.Name()).Add(), Style: excel.EuroStyle()}, 5: excel.Cell{Value: ctx.monthlyOverview.formula(externalCostsChargeable).Reference(ms.refSheet.Name()).Add(), Style: excel.EuroStyle()},
			6: excel.Cell{Value: ctx.monthlyOverview.formula(invoice).Reference(ms.refSheet.Name()).Add(), Style: excel.EuroStyle()},
			7: excel.Cell{Value: ctx.monthlyOverview.formula(honorar).Reference(ms.refSheet.Name()).Add(), Style: excel.EuroStyle()},
		})
	}

	// cells := map[int]excel.Cell{
	// 	1: excel.Cell{Value: ctx.monthlyOverview.formula(customer).Reference(ms.refSheet.Name()).Add(), Style: excel.NoStyle()},
	// 	2: excel.Cell{Value: ms.refSheet., Style: excel.NoStyle()},
	// 	3: excel.Cell{Value: ctx.monthlyOverview.formula(revenue).Reference(ms.refSheet.Name()).Add(), Style: excel.EuroStyle()},
	// 	4: excel.Cell{Value: ctx.monthlyOverview.formula(externalCosts).Reference(ms.refSheet.Name()).Add(), Style: excel.EuroStyle()},
	// 	5: excel.Cell{Value: ctx.monthlyOverview.formula(externalCostsChargeable).Reference(ms.refSheet.Name()).Add(), Style: excel.EuroStyle()},
	// 	6: excel.Cell{Value: ctx.monthlyOverview.formula(invoice).Reference(ms.refSheet.Name()).Add(), Style: excel.EuroStyle()},
	// 	7: excel.Cell{Value: ctx.monthlyOverview.formula(honorar).Reference(ms.refSheet.Name()).Add(), Style: excel.EuroStyle()},
	// }
}
