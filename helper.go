package main

import (
	"fmt"
	"sort"
	"strconv"
)

func mustParseFloat(s string) float32 {
	v, err := strconv.ParseFloat(s, 32)
	if err != nil {
		fmt.Println(err)
		panic("couldn't parse string")
	}
	return float32(v)
}

func jobnrPrefix(jobnr string) string {
	if jobnr == "" {
		return ""
	}
	return jobnr[:4]
}

func campagne(jobnr string) int {
	if jobnr == "" {
		return 0
	}
	value, err := strconv.Atoi(jobnr[5:8])
	if err != nil {
		fmt.Println(err)
	}
	return value
}

func job(jobnr string) int {
	if jobnr == "" {
		return 0
	}
	value, err := strconv.Atoi(jobnr[9:])
	if err != nil {
		fmt.Println(err)
	}
	return value
}

func sum(ofSlice []float32) float32 {
	var sum float32
	for _, f := range ofSlice {
		sum += f
	}
	return sum
}

// sortProjects sorts the provided projects by jobnr
func sortProjects(projects []adjustedProject) []adjustedProject {
	prjs := map[string]adjustedProject{}
	keys := []string{}
	for _, p := range projects {
		prjs[p.jobnr] = p
		keys = append(keys, p.jobnr)
	}

	sort.Strings(keys)
	sortedPrjs := []adjustedProject{}
	for _, key := range keys {
		sortedPrjs = append(sortedPrjs, prjs[key])
	}
	return sortedPrjs
}
