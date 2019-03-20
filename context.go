package main

import (
	"github.com/AnotherCoolDude/excel"
)

type cellMap map[string][]excel.Coordinates

type context struct {
	projectSummary  cellMap
	customerSummary cellMap
	sheetSummary    cellMap
}

func newContext() *context {
	return &context{
		projectSummary:  cellMap{},
		customerSummary: cellMap{},
		sheetSummary:    cellMap{},
	}
}

func (cmap *cellMap) addFromCurrentRow(sh *excel.Sheet, headerList []header) {
	for _, hdr := range headerList {
		(*cmap)[hdr.string()] = append((*cmap)[hdr.string()], excel.Coordinates{Row: sh.CurrentRow(), Column: hdr.int() + 1})
	}
}

func (cmap *cellMap) addFromRow(row int, headerList []header) {
	for _, hdr := range headerList {
		(*cmap)[hdr.string()] = append((*cmap)[hdr.string()], excel.Coordinates{Row: row, Column: hdr.int() + 1})
	}
}

func (cmap *cellMap) formula(hdr header) *excel.Formula {
	return &excel.Formula{Coords: (*cmap)[hdr.string()]}
}
