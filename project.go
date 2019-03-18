package main

import "github.com/AnotherCoolDude/excel"

type project struct {
	customer                string
	jobnr                   string
	revenue                 float32
	externalCosts           float32
	externalCostsChargeable float32
	invoice                 []float32
	activity                []string
	fibu                    []string
	paginiernr              []string
	honorar                 float32
}

func (pjt *project) Columns() []string {
	return headerTitle()
}

func (pjt *project) Insert(sh *excel.Sheet) {
	pjtCells := map[int]excel.Cell{
		jobnr.int():                   excel.Cell{Value: pjt.jobnr, Style: excel.NoStyle()},
		revenue.int():                 excel.Cell{Value: pjt.revenue, Style: excel.EuroStyle()},
		externalCosts.int():           excel.Cell{Value: pjt.externalCosts, Style: excel.EuroStyle()},
		externalCostsChargeable.int(): excel.Cell{Value: pjt.externalCostsChargeable, Style: excel.EuroStyle()},
	}
	sh.AddRow(pjtCells)
	ctx.projectSummary.addFromCurrentRow(sh, []header{jobnr, revenue, externalCosts, externalCostsChargeable})

	erCells := map[int]excel.Cell{}
	for i, inv := range pjt.invoice {
		erCells = map[int]excel.Cell{
			invoice.int():    excel.Cell{Value: inv, Style: excel.EuroStyle()},
			activity.int():   excel.Cell{Value: pjt.activity[i], Style: excel.NoStyle()},
			fibu.int():       excel.Cell{Value: pjt.fibu[i], Style: excel.NoStyle()},
			paginiernr.int(): excel.Cell{Value: pjt.paginiernr[i], Style: excel.NoStyle()},
		}
		sh.AddRow(erCells)
		ctx.projectSummary.addFromCurrentRow(sh, []header{invoice, activity, fibu, paginiernr})
	}
}
