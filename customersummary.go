package main

import (
	"github.com/AnotherCoolDude/excel"
)

type customerSummary struct {
	name string
}

func (csts *customerSummary) Columns() []string {
	return []string{}
}

func (csts *customerSummary) Insert(sh *excel.Sheet) {
	topBorderEuroStyle := excel.Style{Border: excel.Top, Format: excel.Euro}
	topBorderNoStyle := excel.Style{Border: excel.Top, Format: excel.NoFormat}

	cstsCells := map[int]excel.Cell{
		customer.int():                excel.Cell{Value: csts.name, Style: topBorderNoStyle},
		jobnr.int():                   excel.Cell{Value: " ", Style: topBorderNoStyle},
		revenue.int():                 excel.Cell{Value: ctx.customerSummary.formula(revenue).Add(), Style: topBorderEuroStyle},
		externalCosts.int():           excel.Cell{Value: ctx.customerSummary.formula(externalCosts).Add(), Style: topBorderEuroStyle},
		externalCostsChargeable.int(): excel.Cell{Value: ctx.customerSummary.formula(externalCostsChargeable).Add(), Style: topBorderEuroStyle},
		invoice.int():                 excel.Cell{Value: ctx.customerSummary.formula(invoice).Add(), Style: topBorderEuroStyle},
		activity.int():                excel.Cell{Value: " ", Style: topBorderNoStyle},
		fibu.int():                    excel.Cell{Value: " ", Style: topBorderNoStyle},
		paginiernr.int():              excel.Cell{Value: " ", Style: topBorderNoStyle},
		honorar.int():                 excel.Cell{Value: ctx.customerSummary.formula(honorar).Add(), Style: topBorderEuroStyle},
	}
	if len(sh.HeaderColumns()) > 10 {
		cstsCells[subsidiesEL.int()] = excel.Cell{Value: ctx.customerSummary.formula(subsidiesEL).Add(), Style: topBorderEuroStyle}
		cstsCells[subsidiesFK.int()] = excel.Cell{Value: ctx.customerSummary.formula(subsidiesFK).Add(), Style: topBorderEuroStyle}
	}
	sh.AddRow(cstsCells)
	ctx.sheetSummary.addFromCurrentRow(sh, []header{
		revenue,
		externalCosts,
		externalCostsChargeable,
		invoice,
		honorar,
		subsidiesEL,
		subsidiesFK,
	})
	ctx.monthlyOverview.addFromCurrentRow(sh, []header{
		customer,
		revenue,
		externalCosts,
		externalCostsChargeable,
		invoice,
		honorar,
	})
	sh.AddEmptyRow()
	ctx.customerSummary = cellMap{}

}
