package main

import (
	"fmt"
	"log"
)

func main() {
	fmt.Println("Hello World")

	pieces := ParseXLSX("./inputs/RowHeroData.xlsx")

	err := OutputXLSX(pieces, "./outputs/RowHeroData-1.xlsx")

	if err != nil {
		log.Fatal(err)
	}
}
