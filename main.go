package main

import (
	"fmt"
	"strings"

	"github.com/gocolly/colly"
	"github.com/xuri/excelize/v2"
)

type Word struct {
	En string `json:"en"`
	Ru string `json:"ru"`
}

var wordCollection = []Word{}
var cnt int

func main() {
	scrapPage("https://www.en365.ru/top1000.htm")
	scrapPage("https://www.en365.ru/top1000a.htm")
	scrapPage("https://www.en365.ru/top1000b.htm")
	fmt.Printf("cnt %v \n", cnt)
	writeResultXls()
}

func scrapPage(url string) {
	c := colly.NewCollector()

	// Find and visit all links
	c.OnHTML("table tr", func(e *colly.HTMLElement) {

		enWord := e.DOM.Find("td:nth-child(2)").Text()
		ruWord := e.DOM.Find("td:nth-child(3)").Text()
		if !strings.Contains(enWord, "Английское слово") {
			wordCollection = append(wordCollection, Word{En: enWord, Ru: ruWord})
			cnt++
		}
		// fmt.Println("enWord", enWord)
	})

	c.Visit(url)
}

func writeResultXls() {
	xlsx := excelize.NewFile()

	xlsx.NewSheet("Sheet1")

	for i, word := range wordCollection {
		xlsx.SetCellValue("Sheet1", fmt.Sprintf("A%v", i+1), word.En)
		xlsx.SetCellValue("Sheet1", fmt.Sprintf("B%v", i+1), word.Ru)
	}

	// Save spreadsheet by the given path.
	if err := xlsx.SaveAs("./EnRu.xlsx"); err != nil {
		fmt.Println(err)
	}
}
