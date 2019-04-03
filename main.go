package main

import (
	"fmt"

	"github.com/AnotherCoolDude/excel"
)

const (
	// monthly
	rentabilität       = "/Users/christianhovenbitzer/Desktop/Honorar/rent_janfeb.xlsx"
	eingangsrechnungen = "/Users/christianhovenbitzer/Desktop/Honorar/er_novmarch.xlsx"

	// final file
	auswertung = "/Users/christianhovenbitzer/Desktop/Honorar/AuswertungMonth.xlsx"

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

	writeMonthlyEvaluationToFile("Jan Feb")
	//writeYearlyEvaluationToFile()
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

func writeYearlyEvaluationToFile() {
	adj17Excel := excel.File(adj17, "", false)
	adj19Excel := excel.File(adj19, "", false)
	abgr17Excel := excel.File(abgr17, "", false)
	abgr19Excel := excel.File(abgr19, "", false)
	eingangsrechnungen17_19Excel := excel.File(eingangsrechnungen17_19, "", false)
	rentabilität18Excel := excel.File(rentabilität18, "", false)
	projects, adjustments := allocateAdjustedProjects(parseDataForYearlyEvaluation(rentabilität18Excel, eingangsrechnungen17_19Excel, adj17Excel, adj19Excel, abgr17Excel, abgr19Excel))

	auswertungExcel := excel.File(auswertung, "Auswertung 2018", true)
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

	auswertungExcel.Sheet("Adjustments").AddHeaderColumn(adjustmentHeaderTitle())
	for _, adj := range adjustments {
		auswertungExcel.Sheet("Adjustments").Add(&adj)
	}
	auswertungExcel.Sheet("Adjustments").Add(&adjustmentSummary{})
	auswertungExcel.Sheet("Adjustments").FreezeHeader()

	fmt.Printf("writing %d projects to file\n", len(projects)+len(adjustments))
	fmt.Println()
	fmt.Println("saving file...")
	auswertungExcel.Save(auswertung)
}

func writeMonthlyEvaluationToFile(sheetTitle string) {
	rentExcel := excel.File(rentabilität, "", false)
	erExcel := excel.File(eingangsrechnungen, "", false)
	abgr18Excel := excel.File(abgr19, "", false)
	adj18Excel := excel.File(adj19, "", false)

	rentData := rentExcel.FirstSheet().ExtractColumns([]string{
		"A", "C", "E", "G", "I", "L", "E",
	})

	erData := erExcel.FirstSheet().ExtractColumns([]string{
		"A", "F", "G", "I", "K",
	})

	abgr18Data := abgr18Excel.Sheet("konsolidiert").ExtractColumns([]string{
		"A",
	})

	adj18Data := adj18Excel.Sheet("konsolidiert18").ExtractColumns([]string{
		"A", "B", "C",
	})

	projects := allocateAdjustedProjects19(rentData, erData, abgr18Data, adj18Data)

	auswertungExcel := excel.File(auswertung, sheetTitle, true)
	fmt.Printf("writing %d projects to file\n", len(projects))

	// gesamt
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

	auswertungExcel.Sheet("Zusammenfassung").AddHeaderColumn(monthlyOverViewTitle())
	auswertungExcel.Sheet("Zusammenfassung").Add(&monthlyOverview{refSheet: auswertungExcel.FirstSheet()})
	auswertungExcel.Sheet("Zusammenfassung").FreezeHeader()

	// PR
	filtered := filterAdjProjects(projects, func(prj adjustedProject) bool {
		return prj.jobnr[5:6] == "2"
	})
	auswertungExcel.Sheet("Jan Feb PR").AddHeaderColumn(headerTitleSubsidies())
	for i, prj := range filtered {
		auswertungExcel.Sheet("Jan Feb PR").Add(&prj)
		auswertungExcel.Sheet("Jan Feb PR").Add(&projectSummary{})
		if i < len(filtered)-1 && jobnrPrefix(filtered[i+1].jobnr) != jobnrPrefix(prj.jobnr) {
			auswertungExcel.Sheet("Jan Feb PR").Add(&customerSummary{name: prj.customer})
		}
	}
	auswertungExcel.Sheet("Jan Feb PR").Add(&customerSummary{filtered[len(filtered)-1].customer})
	auswertungExcel.Sheet("Jan Feb PR").Add(&sheetSummary{})
	auswertungExcel.Sheet("Jan Feb PR").FreezeHeader()

	auswertungExcel.Sheet("Zusammenfassung PR").AddHeaderColumn(monthlyOverViewTitle())
	auswertungExcel.Sheet("Zusammenfassung PR").Add(&monthlyOverview{refSheet: auswertungExcel.Sheet("Jan Feb PR")})
	auswertungExcel.Sheet("Zusammenfassung PR").FreezeHeader()

	fmt.Println()
	fmt.Println("saving file...")
	auswertungExcel.Save(auswertung)
}

func filterProjects(projects []project, fn func(prj project) bool) {
	b := projects[:0]
	for _, prj := range projects {
		if fn(prj) {
			b = append(b, prj)
		}
	}
	for i := len(b); i < len(projects); i++ {
		projects[i] = project{}
	}
}

func filterAdjProjects(projects []adjustedProject, fn func(prj adjustedProject) bool) []adjustedProject {
	filtered := []adjustedProject{}
	for _, prj := range projects {
		if fn(prj) {
			filtered = append(filtered, prj)
		}
	}
	return filtered
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

func allocateAdjustedProjects19(rentData, erData, abgr18Data, adj18Data [][]string) []adjustedProject {
	abgrProjectnrs := []string{}
	for _, row := range abgr18Data {
		abgrProjectnrs = append(abgrProjectnrs, row[0])
	}

	// cleaned from 18 projects
	rentWithout18 := [][]string{}
	for _, rentRow := range rentData {
		if !contains(abgrProjectnrs, rentRow[2]) {
			rentWithout18 = append(rentWithout18, rentRow)
		}
	}
	adjustments := []adjustment{}
	for _, row := range adj18Data {
		adjustments = append(adjustments, adjustment{
			jobnr:    row[0],
			year:     "2018",
			amountEL: mustParseFloat(row[1]),
			amountFK: mustParseFloat(row[2]),
		})
	}

	adjustedProjects := []adjustedProject{}
	for _, row := range rentWithout18 {
		adjustedProjects = append(adjustedProjects, adjustedProject{
			customer:                row[0],
			jobnr:                   row[1],
			revenue:                 mustParseFloat(row[2]),
			externalCosts:           mustParseFloat(row[4]),
			externalCostsChargeable: mustParseFloat(row[3]),
			invoice:                 []float32{},
			activity:                []string{},
			fibu:                    []string{},
			paginiernr:              []string{},
			honorar:                 0.0,
			subsidiesYear:           []string{},
			subsidiesEL:             []float32{},
			subsidiesFK:             []float32{},
		})
	}
	fmt.Println(len(adjustedProjects))

	for i, adjPrj := range adjustedProjects {
		for _, erRow := range erData {
			if erRow[2] == adjPrj.jobnr {
				adjustedProjects[i].invoice = append(adjustedProjects[i].invoice, mustParseFloat(erRow[4]))
				adjustedProjects[i].activity = append(adjustedProjects[i].activity, erRow[3])
				adjustedProjects[i].fibu = append(adjustedProjects[i].fibu, erRow[1])
				adjustedProjects[i].paginiernr = append(adjustedProjects[i].paginiernr, erRow[0])
			}
		}

		for _, adj := range adjustments {
			if adjPrj.jobnr == adj.jobnr {
				adjustedProjects[i].subsidiesYear = append(adjustedProjects[i].subsidiesYear, adj.year)
				adjustedProjects[i].subsidiesEL = append(adjustedProjects[i].subsidiesEL, adj.amountEL)
				adjustedProjects[i].subsidiesFK = append(adjustedProjects[i].subsidiesFK, adj.amountFK)
			}
		}
	}
	return adjustedProjects

}

func contains(slice []string, item string) bool {
	for _, s := range slice {
		if s == item {
			return true
		}
	}
	return false
}

func allocateAdjustedProjects(rentData, erData, adj17Data, adj19Data, abgr17Data, abgr19Data [][]string) ([]adjustedProject, []adjustment) {
	adjProjects := []adjustedProject{}

	adjustments := []adjustment{}
	for _, row := range adj17Data {
		adjustments = append(adjustments, adjustment{
			jobnr:    row[0],
			year:     "2017",
			amountEL: mustParseFloat(row[1]),
			amountFK: mustParseFloat(row[2]),
		})
	}
	for _, row := range adj19Data {
		adjustments = append(adjustments, adjustment{
			jobnr:    row[0],
			year:     "2019",
			amountEL: mustParseFloat(row[1]),
			amountFK: mustParseFloat(row[2]),
		})
	}

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
	return sortProjects(adjProjects), adjustments
}
