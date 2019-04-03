package main

import (
	"github.com/AnotherCoolDude/excel"
)

type adjustedProject struct {
	customer                string
	jobnr                   string
	revenue                 float32
	externalCosts           float32
	externalCostsChargeable float32
	invoice                 []float32
	activity                []string
	fibu                    []string
	paginiernr              []string
	subsidiesYear           []string
	subsidiesEL             []float32
	subsidiesFK             []float32
	honorar                 float32
}

func (adjpjt *adjustedProject) Columns() []string {
	return headerTitleSubsidies()
}

func (adjpjt *adjustedProject) Insert(sh *excel.Sheet) {
	adjpjtCells := map[int]excel.Cell{
		jobnr.int():                   excel.Cell{Value: adjpjt.jobnr, Style: excel.NoStyle()},
		revenue.int():                 excel.Cell{Value: adjpjt.revenue, Style: excel.EuroStyle()},
		externalCosts.int():           excel.Cell{Value: adjpjt.externalCosts, Style: excel.EuroStyle()},
		externalCostsChargeable.int(): excel.Cell{Value: adjpjt.externalCostsChargeable, Style: excel.EuroStyle()},
	}
	sh.AddRow(adjpjtCells)
	ctx.projectSummary.addFromCurrentRow(sh, []header{jobnr, revenue, externalCosts, externalCostsChargeable})

	cells := map[int]excel.Cell{}
	for i, inv := range adjpjt.invoice {
		cells = map[int]excel.Cell{
			invoice.int():    excel.Cell{Value: inv, Style: excel.EuroStyle()},
			activity.int():   excel.Cell{Value: adjpjt.activity[i], Style: excel.NoStyle()},
			fibu.int():       excel.Cell{Value: adjpjt.fibu[i], Style: excel.NoStyle()},
			paginiernr.int(): excel.Cell{Value: adjpjt.paginiernr[i], Style: excel.NoStyle()},
		}
		sh.AddRow(cells)
		ctx.projectSummary.addFromCurrentRow(sh, []header{invoice})
	}
	// get lenght of the longer subsidies slice
	lenghtSubsidies := len(adjpjt.subsidiesFK)
	if len(adjpjt.subsidiesEL) > len(adjpjt.subsidiesFK) {
		lenghtSubsidies = len(adjpjt.subsidiesEL)
	}

	for i := 0; i < lenghtSubsidies; i++ {
		cells = map[int]excel.Cell{
			13:                excel.Cell{Value: adjpjt.subsidiesYear[i], Style: excel.NoStyle()},
			subsidiesEL.int(): excel.Cell{Value: adjpjt.subsidiesEL[i], Style: excel.EuroStyle()},
			subsidiesFK.int(): excel.Cell{Value: adjpjt.subsidiesFK[i], Style: excel.EuroStyle()},
		}
		sh.AddRow(cells)
		ctx.projectSummary.addFromCurrentRow(sh, []header{subsidiesEL, subsidiesFK})
	}

}
