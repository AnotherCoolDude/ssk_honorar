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
		erFiBu := []string{}
		erPagNr := []string{}
		erInv := []float32{}
		erLA := []string{}
		for _, erRow := range erData {
			//rentRow[2] = jobnr; erRow[6] = jobnr
			if rentRow[2] == erRow[2] {
				erFiBu = append(erFiBu, erRow[1])
				erPagNr = append(erPagNr, erRow[0])
				erInv = append(erInv, mustParseFloat(erRow[4]))
				erLA = append(erLA, erRow[3])
			}
		}
		var inv float32
		for _, i := range erInv {
			inv += i
		}
		projects = append(projects, project{
			customer:                rentRow[0],
			jobnr:                   rentRow[1],
			revenue:                 mustParseFloat(rentRow[2]),
			externalCosts:           mustParseFloat(rentRow[4]),
			externalCostsChargeable: mustParseFloat(rentRow[3]),
			invoice:                 erInv,
			activity:                erLA,
			fibu:                    erFiBu,
			paginiernr:              erPagNr,
			honorar:                 mustParseFloat(rentRow[2]) - inv,
		})
	}

	fmt.Printf("writing %d projects to file\n", len(projects))
	bar := progressbar.New(len(projects))
	//var prevProject project
	for i, prj := range projects {
		auswertungExcel.FirstSheet().Add(&prj)
		bar.Add(1)
		auswertungExcel.FirstSheet().Add(&projectSummary{})
		if i < len(projects)-1 && jobnrPrefix(projects[i+1].jobnr) != jobnrPrefix(prj.jobnr) {
			auswertungExcel.FirstSheet().Add(&customerSummary{name: prj.customer})
		}
		/*
			if jobnrPrefix(prevProject.jobnr) != " " && jobnrPrefix(prj.jobnr) != jobnrPrefix(prevProject.jobnr) {
				auswertungExcel.FirstSheet().Add(&customerSummary{name: prevProject.customer})
			}
		prevProject = prj*/
	}
	auswertungExcel.FirstSheet().Add(&sheetSummary{})
	auswertungExcel.FirstSheet().FreezeHeader()
	fmt.Println()
	fmt.Println("saving file...")
	auswertungExcel.Save(auswertung)
}
