package main

// import (
// 	"fmt"

// 	"github.com/AnotherCoolDude/excel"
// )

// type monthlyOverview struct {
// 	refSheet *excel.Sheet
// }

// func (ms *monthlyOverview) Insert(sh *excel.Sheet) {
// 	custAbbr := ""

// 	for _, row := range ms.refSheet.Draft()[1 : len(ms.refSheet.Draft())-1] {

// 		if len(row) > 1 && row[1].HasValue() {
// 			jobnr := fmt.Sprintf("%s", row[1].Value)
// 			custAbbr = jobnr[:4]
// 		}

// 		if row[0].HasValue() {
// 			sh.AddRow(map[int]excel.Cell{
// 				1: *row[customer.int()-1].ChangeStyle(excel.NoStyle()),
// 				2: excel.Cell{Value: custAbbr, Style: excel.NoStyle()},
// 				3: excel.Cell{Value: "=" + row[revenue.int()-1].Coordinates().StringWithReference(ms.refSheet.Name()), Style: excel.EuroStyle()},
// 				4: excel.Cell{Value: "=" + row[externalCosts.int()-1].Coordinates().StringWithReference(ms.refSheet.Name()), Style: excel.EuroStyle()},
// 				5: excel.Cell{Value: "=" + row[externalCostsChargeable.int()-1].Coordinates().StringWithReference(ms.refSheet.Name()), Style: excel.EuroStyle()},
// 				6: excel.Cell{Value: "=" + row[invoice.int()-1].Coordinates().StringWithReference(ms.refSheet.Name()), Style: excel.EuroStyle()},
// 				7: excel.Cell{Value: "=" + row[subsidiesEL.int()-1].Coordinates().StringWithReference(ms.refSheet.Name()), Style: excel.EuroStyle()},
// 				8: excel.Cell{Value: "=" + row[subsidiesFK.int()-1].Coordinates().StringWithReference(ms.refSheet.Name()), Style: excel.EuroStyle()},
// 				9: excel.Cell{Value: "=" + row[honorar.int()-1].Coordinates().StringWithReference(ms.refSheet.Name()), Style: excel.EuroStyle()},
// 			})
// 			custAbbr = ""
// 		}
// 	}
// 	length := len(sh.Draft())
// 	sh.AddRow(map[int]excel.Cell{
// 		1: excel.Cell{Value: "Gesamt", Style: excel.Style{Border: excel.Top, Format: excel.NoFormat}},
// 		2: excel.Cell{Value: " ", Style: excel.Style{Border: excel.Top, Format: excel.NoFormat}},
// 		3: excel.Cell{Value: excel.FormulaFromRange(excel.Coordinates{Row: 2, Column: 3}, excel.Coordinates{Row: length, Column: 3}).Sum(), Style: excel.Style{Border: excel.Top, Format: excel.Euro}},
// 		4: excel.Cell{Value: excel.FormulaFromRange(excel.Coordinates{Row: 2, Column: 4}, excel.Coordinates{Row: length, Column: 4}).Sum(), Style: excel.Style{Border: excel.Top, Format: excel.Euro}},
// 		5: excel.Cell{Value: excel.FormulaFromRange(excel.Coordinates{Row: 2, Column: 5}, excel.Coordinates{Row: length, Column: 5}).Sum(), Style: excel.Style{Border: excel.Top, Format: excel.Euro}},
// 		6: excel.Cell{Value: excel.FormulaFromRange(excel.Coordinates{Row: 2, Column: 6}, excel.Coordinates{Row: length, Column: 6}).Sum(), Style: excel.Style{Border: excel.Top, Format: excel.Euro}},
// 		7: excel.Cell{Value: excel.FormulaFromRange(excel.Coordinates{Row: 2, Column: 7}, excel.Coordinates{Row: length, Column: 7}).Sum(), Style: excel.Style{Border: excel.Top, Format: excel.Euro}},
// 		8: excel.Cell{Value: excel.FormulaFromRange(excel.Coordinates{Row: 2, Column: 8}, excel.Coordinates{Row: length, Column: 8}).Sum(), Style: excel.Style{Border: excel.Top, Format: excel.Euro}},
// 		9: excel.Cell{Value: excel.FormulaFromRange(excel.Coordinates{Row: 2, Column: 9}, excel.Coordinates{Row: length, Column: 9}).Sum(), Style: excel.Style{Border: excel.Top, Format: excel.Euro}},
// 	})
// }
