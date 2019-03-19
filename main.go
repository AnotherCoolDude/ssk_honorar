package main

import (
	"fmt"
	"github.com/AnotherCoolDude/excel"
	"github.com/schollz/progressbar"
)

const (
	rentabilität       = "/Users/christianhovenbitzer/Desktop/Honorar/rent_janfeb.xlsx"
	eingangsrechnungen = "/Users/christianhovenbitzer/Desktop/Honorar/er_novmarch.xlsx"
	auswertung         = "/Users/christianhovenbitzer/Desktop/Honorar/result.xlsx"
)

var ctx *context

func main() {
	ctx = newContext()

	rentExcel := excel.File(rentabilität, "")
	erExcel := excel.File(eingangsrechnungen, "")

	auswertungExcel := excel.File(auswertung, "jan feb")

	rentData := rentExcel.FirstSheet().FilterByColumn([]string{
		"A", "C", "E", "G", "I", "L", "E",
	})

	erData := erExcel.FirstSheet().FilterByColumn([]string{
		"A", "F", "G", "I", "K",
	})

	projects := []project{}
	for _, rentRow := range rentData {
		newPrj := project{
			customer:                rentRow[0],
			jobnr:                   rentRow[1],
			revenue:                 mustParseFloat(rentRow[2]),
			externalCosts:           mustParseFloat(rentRow[4]),
			externalCostsChargeable: mustParseFloat(rentRow[3]),
			invoice:                 []float32{},
			activity:                []string{},
			fibu:                    []string{},
			paginiernr:              []string{},
			honorar:                 0.0,
		}
		for _, erRow := range erData {
			if erRow[2] == newPrj.jobnr {
				newPrj.invoice = append(newPrj.invoice, mustParseFloat(erRow[4]))
				newPrj.activity = append(newPrj.activity, erRow[3])
				newPrj.fibu = append(newPrj.fibu, erRow[1])
				newPrj.paginiernr = append(newPrj.paginiernr, erRow[0])
			}
		}
		projects = append(projects, newPrj)
	}

	fmt.Printf("writing %d projects to file\n", len(projects))
	bar := progressbar.New(len(projects))

	for i, prj := range projects {
		auswertungExcel.FirstSheet().Add(&prj)
		bar.Add(1)
		auswertungExcel.FirstSheet().Add(&projectSummary{})
		if i < len(projects)-1 && jobnrPrefix(projects[i+1].jobnr) != jobnrPrefix(prj.jobnr) {
			auswertungExcel.FirstSheet().Add(&customerSummary{name: prj.customer})
		}
	}
	auswertungExcel.FirstSheet().Add(&sheetSummary{})
	auswertungExcel.FirstSheet().FreezeHeader()
	fmt.Println()
	fmt.Println("saving file...")
	auswertungExcel.Save(auswertung)
}
