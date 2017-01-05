package main

import (
	"fmt"
	"io/ioutil"
	"os"
	"path"

	"github.com/tealeg/xlsx"
)

const (
	A = iota
	B
	C
	D
	E
	F
	G
	H
	I
	J
	K
	L
	M
	N
	O
	P
	Q
	R
	S
	T
	U
	V
	W
	X
	Y
	Z
)

const (
	SIG = "本软件由福建工程学院白凤军开发，用于打印名册。"
)

func main() {
	fmt.Println(SIG)
	if len(os.Args) != 2 {
		fmt.Printf("usage: %s %s\n", path.Base(os.Args[0]), "file.xlsx")
		return
	}

	var (
		xlfile *xlsx.File
		sheet  *xlsx.Sheet
		err    error
	)
	if xlfile, err = xlsx.OpenFile(os.Args[1]); err != nil {
		fmt.Println("Can't openfile: ", os.Args[1])
		return
	}
	body := "%%"
	sheet = xlfile.Sheets[0]
	var i int = 0
	const (
		spaceY     = 6.5
		textOffset = 1.0
	)
	linesOfPage := (len(sheet.Rows)-6)/2 - 2
	fmt.Printf("lines of page: %d\n", linesOfPage)
	for rowsLen := len(sheet.Rows); rowsLen >= 6; rowsLen-- {
		text, _ := sheet.Cell(rowsLen-1, C).String()
		if text == "" {
			continue
		}
		j := i / linesOfPage
		x, y := float64(j)*40.0+10.0, float64(i-j*linesOfPage)*spaceY
		tx, ty := float64(j)*40.0+10.0, float64(i-j*linesOfPage)*spaceY+textOffset
		itemText := fmt.Sprintf(drawFormat, x, y, tx, ty, text)
		body = fmt.Sprintf("%s%s", body, itemText)
		i++
	}
	body = fmt.Sprintf(fmtStr, body)
	ioutil.WriteFile("x.tex", []byte(body), 0777)
	fmt.Printf("%s\n", "generate x.tex file.")
}

const (
	fmtStr = `
	\documentclass[10pt]{ctexart}
	\usepackage{fancyhdr}
	\usepackage[a4paper,hmargin={1cm,1cm},vmargin={1cm,1cm}]{geometry}
	\begin{document}
	\setlength{\unitlength}{1.0mm}
	\begin{picture}(20,260)
	%s
	\end{picture}
	\end{document}
	`
	drawFormat = `
	\put(%f, %f){\line(1,0){27}}
	\put(%f, %f) {%s}
	`
)
