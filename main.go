package main

import (
	"fmt"
	"github.com/AnotherCoolDude/excel"
)

const (
	// monthly
	rentabilität       = "/Users/christianhovenbitzer/Desktop/Honorar/rent_janfeb.xlsx"
	eingangsrechnungen = "/Users/christianhovenbitzer/Desktop/Honorar/er_novmarch.xlsx"
	auswertung         = "/Users/christianhovenbitzer/Desktop/Honorar/result_adjusted.xlsx"

	// adjusted yearly
	adj17                   = "/Users/christianhovenbitzer/Desktop/Honorar/2018/adjusted17.xlsx"
	adj19                   = "/Users/christianhovenbitzer/Desktop/Honorar/2018/adjusted19.xlsx"
	abgr17                  = "/Users/christianhovenbitzer/Desktop/Honorar/2018/abgrenzung17.xlsx"
	abgr19                  = "/Users/christianhovenbitzer/Desktop/Honorar/2018/abgrenzung19.xlsx"
	eingangsrechnungen17_19 = "/Users/christianhovenbitzer/Desktop/Honorar/2018/er_rechnungsbuch_17-19.xlsx"
	rentabilität18          = "/Users/christianhovenbitzer/Desktop/Honorar/2018/rentabilität18.xlsx"
)

var ctx *context

func main() {
	ctx = newContext()

	// rentExcel := excel.File(rentabilität, "")
	// erExcel := excel.File(eingangsrechnungen, "")
	// projects := allocateProjects(parseDataForMonthlyEvaluation(rentExcel, erExcel))

	adj17Excel := excel.File(adj17, "")
	adj19Excel := excel.File(adj19, "")
	abgr17Excel := excel.File(abgr17, "")
	abgr19Excel := excel.File(abgr19, "")
	eingangsrechnungen17_19Excel := excel.File(eingangsrechnungen17_19, "")
	rentabilität18Excel := excel.File(rentabilität18, "")
	projects := allocateAdjustedProjects(parseDataForYearlyEvaluation(rentabilität18Excel, eingangsrechnungen17_19Excel, adj17Excel, adj19Excel, abgr17Excel, abgr19Excel))

	auswertungExcel := excel.File(auswertung, "jan feb")
	fmt.Printf("writing %d projects to file\n", len(projects))
	auswertungExcel.FirstSheet().AddHeaderColumn(headerTitleSubsidies())
	for i, prj := range projects {
		auswertungExcel.FirstSheet().Add(&prj)
		auswertungExcel.FirstSheet().Add(&projectSummary{})
		if i < len(projects)-1 && jobnrPrefix(projects[i+1].jobnr) != jobnrPrefix(prj.jobnr) {
			auswertungExcel.FirstSheet().Add(&customerSummary{name: prj.customer})
		}
	}
	auswertungExcel.FirstSheet().Add(&customerSummary{projects[len(projects)-1].customer})

	auswertungExcel.FirstSheet().Add(&sheetSummary{})
	auswertungExcel.FirstSheet().FreezeHeader()
	fmt.Println()
	fmt.Println("saving file...")
	auswertungExcel.Save(auswertung)
}

func parseDataForMonthlyEvaluation(rentFile, erFile *excel.Excel) (rentData, erData [][]string) {
	rentData = rentFile.FirstSheet().ExtractColumns([]string{
		"A", "C", "E", "G", "I", "L", "E",
	})

	erData = erFile.FirstSheet().ExtractColumns([]string{
		"A", "F", "G", "I", "K",
	})
	return
}

func parseDataForYearlyEvaluation(rentFile, erFile, adj17File, adj19File, abgr17File, abgr19File *excel.Excel) (rentData, erData, adj17Data, adj19Data, abgr17Data, abgr19Data [][]string) {
	rentData = rentFile.FirstSheet().ExtractColumns([]string{
		"A", "C", "E", "G", "I", "L", "E",
	})

	erData = erFile.FirstSheet().ExtractColumns([]string{
		"A", "F", "G", "I", "K",
	})

	konsolidiert := "konsolidiert"

	adj17Data = adj17File.Sheet(konsolidiert).ExtractColumns([]string{
		"A", "B", "C",
	})
	adj19Data = adj19File.Sheet(konsolidiert).ExtractColumns([]string{
		"A", "B", "C",
	})
	abgr17Data = abgr17File.Sheet(konsolidiert).ExtractColumns([]string{
		"A",
	})
	abgr19Data = abgr19File.Sheet(konsolidiert).ExtractColumns([]string{
		"A", "B", "C",
	})
	return
}

func allocateProjects(rentData, erData [][]string) []project {
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
	return projects
}

func allocateAdjustedProjects(rentData, erData, adj17Data, adj19Data, abgr17Data, abgr19Data [][]string) []adjustedProject {
	adjProjects := []adjustedProject{}

	// add projects from 19
	for _, abgr19Row := range abgr19Data {
		newRentDataRow := []string{abgr19Row[1], abgr19Row[0], abgr19Row[2], "0", "0"}
		rentData = append(rentData, newRentDataRow)
	}

	for _, rentRow := range rentData {
		found := false
		// prevent projects from abgr17Data to be added
		for _, abgr17Row := range abgr17Data {
			if abgr17Row[0] == rentRow[1] {
				found = true
			}
		}
		if found {
			continue
		}
		// create new Project
		newPrj := adjustedProject{
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
			subsidiesYear:           []string{},
			subsidiesEL:             []float32{},
			subsidiesFK:             []float32{},
		}
		for _, erRow := range erData {
			if erRow[2] == newPrj.jobnr {
				newPrj.invoice = append(newPrj.invoice, mustParseFloat(erRow[4]))
				newPrj.activity = append(newPrj.activity, erRow[3])
				newPrj.fibu = append(newPrj.fibu, erRow[1])
				newPrj.paginiernr = append(newPrj.paginiernr, erRow[0])
			}
		}
		for _, adj17Row := range adj17Data {
			if adj17Row[0] == newPrj.jobnr {
				newPrj.subsidiesYear = append(newPrj.subsidiesYear, "Anteil 2017")
				newPrj.subsidiesEL = append(newPrj.subsidiesEL, mustParseFloat(adj17Row[1]))
				newPrj.subsidiesFK = append(newPrj.subsidiesFK, mustParseFloat(adj17Row[2]))
			}
		}
		for _, adj19Row := range adj19Data {
			if adj19Row[0] == newPrj.jobnr {
				newPrj.subsidiesYear = append(newPrj.subsidiesYear, "Anteil 2019")
				newPrj.subsidiesEL = append(newPrj.subsidiesEL, mustParseFloat(adj19Row[1]))
				// adj19Row[2] = Anteil FK 18
				newPrj.subsidiesFK = append(newPrj.subsidiesFK, sum(newPrj.invoice)-mustParseFloat(adj19Row[2]))
			}
		}
		adjProjects = append(adjProjects, newPrj)
	}
	return sortProjects(adjProjects)
}
