package main

import (
	"fmt"
	"github.com/xuri/excelize/v2"
	"log"
)

func main() {

	arquivo := "Plano Mensal_VG.xlsx"

	f, err := excelize.OpenFile(arquivo)
	if err != nil {
		log.Fatalln(err)
	}

	fmt.Println("Nome do arquivo:",arquivo)

	// Realiza validação da planilha e havendo erro, devolve para o usuário corrigir...
	validacao(f)

	// Lê os dados da planilha e imprime na console...
	report(f)

}

func validacao(f *excelize.File) {

	// Abre a primeira pasta (Sheet)
	sheet := f.WorkBook.Sheets.Sheet[0].Name

	// Se a aba estiver com nome errado, já gera um erro
	if sheet != "IS_monthly Plan_GL"  {
		log.Fatalln("ERRO -> Aba IS_monthly Plan_GL não encontrada!")
	}

	cell, err := f.GetCellValue(sheet, "A8")
	if err != nil {
		fmt.Println(err)
		return
	}

	// Verifica os nomes fixos na planilha...
	if cell != "Premium" {
		log.Fatalln("ERRO -> Célula A8 deveria conter a descrição Premium")
	}

	// fmt.Println("Valor da célula A8: ",cell)

	cell, err = f.GetCellValue(sheet, "B8")
	if err != nil {
		fmt.Println(err)
		return
	}

	// Recupera uma fórmula para validação tb
	//formula, err := f.GetCellFormula(sheet,"B8")
	//fmt.Println("Fórmula da célula B8: ", formula)

	// Verifica se a coluna de mês está na posição correta e se o nome também está correto...
	jan21, err := f.GetCellValue(sheet, "B6")
	if err != nil {
		fmt.Println(err)
		return
	}

//	fmt.Println("Valor da célula B6: ", jan21)

	if jan21 != "Jan/21" {
		log.Fatalln("ERRO -> A célula B6 deveria ter o título 'Jan/21'")
	}

	// Verifica se o título da planilha está correto
	titulo, err := f.GetCellValue(sheet, "B5")
	if err != nil {
		fmt.Println(err)
		return
	}

//	fmt.Println("Valor da célula B5: ", titulo)

	if titulo != "Plan 2021 - STA" {
		log.Fatalln("ERRO -> O título na célula B6 deve ser 'Plan 2021 - STA'")
	}

}

func report(f *excelize.File) {

	sheet := f.WorkBook.Sheets.Sheet[0].Name

	a8, err := f.GetCellValue(sheet, "A8")
	if err != nil {
		fmt.Println(err)
		return
	}

	b8, err := f.GetCellValue(sheet, "B8")
	if err != nil {
		fmt.Println(err)
		return
	}

	c8, err := f.GetCellValue(sheet, "C8")
	if err != nil {
		fmt.Println(err)
		return
	}

	d8, err := f.GetCellValue(sheet, "D8")
	if err != nil {
		fmt.Println(err)
		return
	}

	a9, err := f.GetCellValue(sheet, "A9")
	if err != nil {
		fmt.Println(err)
		return
	}

	b9, err := f.GetCellValue(sheet, "B9")
	if err != nil {
		fmt.Println(err)
		return
	}

	c9, err := f.GetCellValue(sheet, "C9")
	if err != nil {
		fmt.Println(err)
		return
	}

	d9, err := f.GetCellValue(sheet, "D9")
	if err != nil {
		fmt.Println(err)
		return
	}

	a10, err := f.GetCellValue(sheet, "A10")
	if err != nil {
		fmt.Println(err)
		return
	}

	b10, err := f.GetCellValue(sheet, "B10")
	if err != nil {
		fmt.Println(err)
		return
	}

	c10, err := f.GetCellValue(sheet, "C10")
	if err != nil {
		fmt.Println(err)
		return
	}

	d10, err := f.GetCellValue(sheet, "D10")
	if err != nil {
		fmt.Println(err)
		return
	}

	a14, err := f.GetCellValue(sheet, "A14")
	if err != nil {
		fmt.Println(err)
		return
	}
	a14 = formatacao(a14)

	b14, err := f.GetCellValue(sheet, "B14")
	if err != nil {
		fmt.Println(err)
		return
	}
	b14 = formatacao(b14)

	c14, err := f.GetCellValue(sheet, "C14")
	if err != nil {
		fmt.Println(err)
		return
	}
	c14 = formatacao(c14)

	d14, err := f.GetCellValue(sheet, "D14")
	if err != nil {
		fmt.Println(err)
		return
	}
	d14 = formatacao(d14)

	a16, err := f.GetCellValue(sheet, "A16")
	if err != nil {
		fmt.Println(err)
		return
	}
	a16 = formatacao(a16)

	b16, err := f.GetCellValue(sheet, "B16")
	if err != nil {
		fmt.Println(err)
		return
	}
	b16 = formatacao(b16)

	c16, err := f.GetCellValue(sheet, "C16")
	if err != nil {
		fmt.Println(err)
		return
	}
	c16 = formatacao(c16)

	d16, err := f.GetCellValue(sheet, "D16")
	if err != nil {
		fmt.Println(err)
		return
	}
	d16 = formatacao(d16)

	fmt.Println("=================================================================================================")

	fmt.Printf("%-20s", a8)
	fmt.Printf("%-20s","")
	fmt.Printf("%-10s",b8)
	fmt.Printf("%-10s",c8)
	fmt.Printf("%-10s",d8)
	fmt.Println()

	fmt.Printf("%-20s", "Premium")
	fmt.Printf("%-20s",a9)
	fmt.Printf("%-10s",b9)
	fmt.Printf("%-10s",c9)
	fmt.Printf("%-10s",d9)
	fmt.Println()

	fmt.Printf("%-20s", "Premium")
	fmt.Printf("%-20s", a10)
	fmt.Printf("%-10s",b10)
	fmt.Printf("%-10s",c10)
	fmt.Printf("%-10s",d10)
	fmt.Println()

	fmt.Printf("%-20s", a14)
	fmt.Printf("%-20s","")
	fmt.Printf("%-10s",b14)
	fmt.Printf("%-10s",c14)
	fmt.Printf("%-10s",d14)
	fmt.Println()

	fmt.Printf("%-20s", a16)
	fmt.Printf("%-20s","")
	fmt.Printf("%-10s",b16)
	fmt.Printf("%-10s",c16)
	fmt.Printf("%-10s",d16)
	fmt.Println()

	fmt.Println("=================================================================================================")

}

func formatacao(valor string) string {
	if valor == "-" {
		return "0"
	} else {
		return valor
	}
}





/*  Importações
go get github.com/xuri/excelize/v2

 */

