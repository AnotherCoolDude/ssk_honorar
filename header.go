package main

import (
	"fmt"
)

const (
	customer header = iota
	jobnr
	revenue
	externalCosts
	externalCostsChargeable
	invoice
	activity
	fibu
	paginiernr
	honorar
)

type header int

func headerTitle() []string {
	return []string{
		"Kunde", "Jobnr", "Umsatz", "FK nwb", "FK wb", "Eingangsrechnungen", "Leistungsart", "FiBu", "Paginiernr.", "Honorar (DB1)",
	}
}

func (hdr header) int() int {
	return int(hdr)
}

func (hdr header) string() string {
	switch hdr {
	case customer:
		return "Kunde"
	case jobnr:
		return "Jobnr"
	case revenue:
		return "Umsatz"
	case externalCosts:
		return "FK nwb"
	case externalCostsChargeable:
		return "FK wb"
	case invoice:
		return "Eingangsrechnungen"
	case activity:
		return "Leistungsart"
	case fibu:
		return "FiBu"
	case paginiernr:
		return "Paginiernr."
	case honorar:
		return "Honorar (DB1)"
	default:
		fmt.Print("Unknown header")
		return " "
	}
}
