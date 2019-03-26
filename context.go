package main

import (
	"github.com/AnotherCoolDude/excel"
)

type cellMap map[string][]excel.Coordinates

type context struct {
	projectSummary    cellMap
	customerSummary   cellMap
	sheetSummary      cellMap
	adjustmentSummary cellMap
	monthlyOverview   cellMap
}

func newContext() *context {
	return &context{
		projectSummary:    cellMap{},
		customerSummary:   cellMap{},
		sheetSummary:      cellMap{},
		adjustmentSummary: cellMap{},
		monthlyOverview:   cellMap{},
	}
}

func (cmap *cellMap) addFromCurrentRow(sh *excel.Sheet, headerList []header) {
	for _, hdr := range headerList {
		(*cmap)[hdr.string()] = append((*cmap)[hdr.string()], excel.Coordinates{Row: sh.CurrentRow(), Column: hdr.int()})
	}
}

func (cmap *cellMap) addFromRow(row int, headerList []header) {
	for _, hdr := range headerList {
		(*cmap)[hdr.string()] = append((*cmap)[hdr.string()], excel.Coordinates{Row: row, Column: hdr.int()})
	}
}

func (cmap *cellMap) formula(hdr header) *excel.Formula {
	coords := (*cmap)[hdr.string()]
	return &excel.Formula{Coords: &coords}
}
