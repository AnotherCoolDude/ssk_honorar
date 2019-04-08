package main

// import "github.com/AnotherCoolDude/excel"

// type sheetSummary struct {
// }

// func (shs *sheetSummary) Columns() []string {
// 	return []string{}
// }

// func (shs *sheetSummary) Insert(sh *excel.Sheet) {

// 	topBorderEuroStyle := excel.Style{Border: excel.Top, Format: excel.Euro}
// 	topBorderNoStyle := excel.Style{Border: excel.Top, Format: excel.NoFormat}

// 	shsCells := map[int]excel.Cell{
// 		customer.int():                excel.Cell{Value: "Gesamt", Style: topBorderNoStyle},
// 		jobnr.int():                   excel.Cell{Value: " ", Style: topBorderNoStyle},
// 		revenue.int():                 excel.Cell{Value: ctx.sheetSummary.formula(revenue).Add(), Style: topBorderEuroStyle},
// 		externalCosts.int():           excel.Cell{Value: ctx.sheetSummary.formula(externalCosts).Add(), Style: topBorderEuroStyle},
// 		externalCostsChargeable.int(): excel.Cell{Value: ctx.sheetSummary.formula(externalCostsChargeable).Add(), Style: topBorderEuroStyle},
// 		invoice.int():                 excel.Cell{Value: ctx.sheetSummary.formula(invoice).Add(), Style: topBorderEuroStyle},
// 		activity.int():                excel.Cell{Value: " ", Style: topBorderNoStyle},
// 		fibu.int():                    excel.Cell{Value: " ", Style: topBorderNoStyle},
// 		paginiernr.int():              excel.Cell{Value: " ", Style: topBorderNoStyle},
// 		honorar.int():                 excel.Cell{Value: ctx.sheetSummary.formula(honorar).Add(), Style: topBorderEuroStyle},
// 	}
// 	if len(sh.HeaderColumns()) > 10 {
// 		shsCells[subsidiesEL.int()] = excel.Cell{Value: ctx.sheetSummary.formula(subsidiesEL).Add(), Style: topBorderEuroStyle}
// 		shsCells[subsidiesFK.int()] = excel.Cell{Value: ctx.sheetSummary.formula(subsidiesFK).Add(), Style: topBorderEuroStyle}
// 	}

// 	sh.AddEmptyRow()
// 	sh.AddRow(shsCells)
// 	ctx.sheetSummary = cellMap{}
// }
