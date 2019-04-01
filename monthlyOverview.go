package main

import (
	"fmt"

	"github.com/AnotherCoolDude/excel"
)

type monthlyOverview struct {
	refSheet *excel.Sheet
	prOnly   bool
}

func (ms *monthlyOverview) Insert(sh *excel.Sheet) {
	abbr := ""
	for _, row := range ms.refSheet.Draft()[1:] {

		if ms.prOnly && fmt.Sprintf("%s", row[1].Value)[5:6] != "2" {
			continue
		}

		if len(row) > 1 && row[1].HasValue() {
			abbr = fmt.Sprintf("%s", row[1].Value)[:4]
		}
		if row[0].Value != excel.DraftCell {
			sh.AddRow(map[int]excel.Cell{
				1: *row[customer.int()].ChangeStyle(excel.NoStyle()),
				2: excel.Cell{Value: abbr, Style: excel.NoStyle()},
				3: excel.Cell{Value: "=" + row[revenue.int()].Coordinates().StringWithReference(ms.refSheet.Name()), Style: excel.EuroStyle()},
				4: excel.Cell{Value: "=" + row[externalCosts.int()].Coordinates().StringWithReference(ms.refSheet.Name()), Style: excel.EuroStyle()},
				5: excel.Cell{Value: "=" + row[externalCostsChargeable.int()].Coordinates().StringWithReference(ms.refSheet.Name()), Style: excel.EuroStyle()},
				6: excel.Cell{Value: "=" + row[invoice.int()].Coordinates().StringWithReference(ms.refSheet.Name()), Style: excel.EuroStyle()},
				7: excel.Cell{Value: "=" + row[subsidiesEL.int()].Coordinates().StringWithReference(ms.refSheet.Name()), Style: excel.EuroStyle()},
				8: excel.Cell{Value: "=" + row[subsidiesFK.int()].Coordinates().StringWithReference(ms.refSheet.Name()), Style: excel.EuroStyle()},
				9: excel.Cell{Value: "=" + row[honorar.int()].Coordinates().StringWithReference(ms.refSheet.Name()), Style: excel.EuroStyle()},
			})
			abbr = ""
		}
	}
}
