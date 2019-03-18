package main

import (
	"fmt"
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

func mustParseInt(s string) int {
	v, err := strconv.Atoi(s)
	if err != nil {
		fmt.Println(err)
		panic("couldn't parse string")
	}
	return v
}

func jobnrPrefix(jobnr string) string {
	if jobnr == "" {
		return ""
	}
	return jobnr[:4]
}
