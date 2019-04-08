package main

import (
	"fmt"
	"github.com/AnotherCoolDude/excel"
)

func writeProjectsOfMonthlyEvaluationToFile(projects []adjustedProject) {
	auswertungExcel := excel.File(auswertung, "Jan Feb", true)
	fmt.Printf("writing %d projects to file\n", len(projects))

	// gesamt
	auswertungExcel.FirstSheet().AddHeaderColumn(headerTitleSubsidies())
	for _, adj := range projects {
		adj.addRowToSheet(auswertungExcel.FirstSheet())
	}
	// for i, prj := range projects {
	// 	auswertungExcel.FirstSheet().Add(&prj)
	// 	auswertungExcel.FirstSheet().Add(&projectSummary{})
	// 	if i < len(projects)-1 && jobnrPrefix(projects[i+1].jobnr) != jobnrPrefix(prj.jobnr) {
	// 		auswertungExcel.FirstSheet().Add(&customerSummary{name: prj.customer})
	// 	}
	// }
	// auswertungExcel.FirstSheet().Add(&customerSummary{projects[len(projects)-1].customer})
	// auswertungExcel.FirstSheet().Add(&sheetSummary{})
	// auswertungExcel.FirstSheet().FreezeHeader()

	// auswertungExcel.Sheet("Zusammenfassung").AddHeaderColumn(monthlyOverViewTitle())
	// auswertungExcel.Sheet("Zusammenfassung").Add(&monthlyOverview{refSheet: auswertungExcel.FirstSheet()})
	// auswertungExcel.Sheet("Zusammenfassung").FreezeHeader()

	// // PR
	// filtered := filterAdjProjects(projects, func(prj adjustedProject) bool {
	// 	return prj.jobnr[5:6] == "2"
	// })
	// auswertungExcel.Sheet("Jan Feb PR").AddHeaderColumn(headerTitleSubsidies())
	// for i, prj := range filtered {
	// 	auswertungExcel.Sheet("Jan Feb PR").Add(&prj)
	// 	auswertungExcel.Sheet("Jan Feb PR").Add(&projectSummary{})
	// 	if i < len(filtered)-1 && jobnrPrefix(filtered[i+1].jobnr) != jobnrPrefix(prj.jobnr) {
	// 		auswertungExcel.Sheet("Jan Feb PR").Add(&customerSummary{name: prj.customer})
	// 	}
	// }
	// auswertungExcel.Sheet("Jan Feb PR").Add(&customerSummary{filtered[len(filtered)-1].customer})
	// auswertungExcel.Sheet("Jan Feb PR").Add(&sheetSummary{})
	// auswertungExcel.Sheet("Jan Feb PR").FreezeHeader()

	// auswertungExcel.Sheet("Zusammenfassung PR").AddHeaderColumn(monthlyOverViewTitle())
	// auswertungExcel.Sheet("Zusammenfassung PR").Add(&monthlyOverview{refSheet: auswertungExcel.Sheet("Jan Feb PR")})
	// auswertungExcel.Sheet("Zusammenfassung PR").FreezeHeader()

	// fmt.Println()
	// fmt.Println("saving file...")
	auswertungExcel.Save(auswertung)
}
