package main

import (
	"fmt"
)

func parse(rentData, erData, abgr18Data, adj18Data [][]string) []adjustedProject {
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
