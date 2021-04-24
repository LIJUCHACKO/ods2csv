// Version-: 05-11-2017

//////////////////////////////////////contents.xml format///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
//<office:spreadsheet>                                                                          							                                      //
//                                                                                                           								                      //
// <table:table table:name="name" table:style-name="ta1">                                                                                                     		                      //
//                                                                                                          								                      //
//      <table:table-row table:number-rows-repeated="2" table:style-name="ro1">                              								                      //
//                                                                                                           								                      //
//      <table:table-cell table:formula="of:=3*[.B2]"  table:number-columns-repeated="2" table:style-name="ce1" office:value-type="string" calcext:value-type="string" office:date-value="" > //
//      <text:p>SrNo<text:span text:style-name="T1">gj</text:span><text:s text:c="10"/><text:s text:c="10"/>gh</text:p>                                                                       //
//      </table:table-cell>                                                                                  								                      //
//                                                                                                                                                  			                      //
// </table:table-row>                                                                                        								                      //
//                                                                                                           								                      //
// </office:spreadsheet>                                                                                     								                      //
////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

package ods

import (
	"archive/zip"
	"fmt"
	"io/ioutil"
	"strconv"
	"strings"

	"github.com/LIJUCHACKO/XmlDB"
)

type Cell struct {
	Type      string //Type float,string ...    ( office:value-type= )
	Value     string //Value                    ( office:value= )
	DateValue string //DateValue                ( office:date-value= )
	Formula   string //formula                  ( table:formula= )
	Text      string //Text

}

type Row struct {
	Cells []Cell
}

type Sheet struct {
	Name string
	Rows []Row
}

type Odsfile struct {
	Sheets []Sheet
}

func ReplaceHTMLSpecialEntities(input string) string {
	output := strings.Replace(input, "&amp;", "&", -1)
	output = strings.Replace(output, "&lt;", "<", -1)
	output = strings.Replace(output, "&gt;", ">", -1)
	output = strings.Replace(output, "&quot;", "\"", -1)
	output = strings.Replace(output, "&lsquo;", "‘", -1)
	output = strings.Replace(output, "&rsquo;", "’", -1)
	output = strings.Replace(output, "&tilde;", "~", -1)
	output = strings.Replace(output, "&ndash;", "–", -1)
	output = strings.Replace(output, "&mdash;", "—", -1)
	output = strings.Replace(output, "&apos;", "'", -1)

	return output
}

func ReadODSFile(odsfilename string) (Odsfile, error) {
	var odsfileContents Odsfile
	var DB *xmlDB.Database = new(xmlDB.Database)
	DB.Debug_enabled = false
	r, err := zip.OpenReader(odsfilename)
	if err != nil {
		return odsfileContents, err
	}
	defer r.Close()

	for _, f := range r.File {
		if f.Name == "content.xml" {
			rc, fileerr1 := f.Open()
			if fileerr1 != nil {
				return odsfileContents, fileerr1
			}
			xmlfile, fileerr := ioutil.ReadAll(rc)
			if fileerr != nil {
				return odsfileContents, fileerr
			}

			xmlline := string(xmlfile)
			xmllines := strings.Split(xmlline, "\n")
			xmlDB.Load_dbcontent(DB, xmllines)
			fmt.Printf("\ncontent.xml loaded to xmldb\n")
			csvSpreadSheets := []Sheet{}
			spreadSheets, _ := xmlDB.GetNode(DB, 0, "office:body/office:spreadsheet/table:table")

			for _, spreadsheet := range spreadSheets {
				tablename := xmlDB.GetNodeAttribute(DB, spreadsheet, "table:name")
				fmt.Printf("\n%s\n", tablename)
				csvrows := []Row{}
				rows, _ := xmlDB.GetNode(DB, spreadsheet, "table:table-row")
				blankrows := 0

				for index, row := range rows {
					fmt.Printf("\r%d/%d", index, len(rows))
					rowisblank := true
					row_repetitionTXT := xmlDB.GetNodeAttribute(DB, row, "table:number-rows-repeated")
					row_repetition := 1
					if len(strings.TrimSpace(row_repetitionTXT)) > 0 {
						value, err := strconv.Atoi(row_repetitionTXT)
						if err != nil {
							return odsfileContents, err
						}
						row_repetition = value
					}
					csvcells := []Cell{}
					Cells, _ := xmlDB.GetNode(DB, row, "table:table-cell")
					blankcells := 0
					for _, cell := range Cells {
						cell_repetitionTXT := xmlDB.GetNodeAttribute(DB, cell, "table:number-columns-repeated")
						cell_repetition := 1
						if len(strings.TrimSpace(cell_repetitionTXT)) > 0 {
							value, err := strconv.Atoi(cell_repetitionTXT)
							if err != nil {
								return odsfileContents, err
							}
							cell_repetition = value
						}
						celltype := xmlDB.GetNodeAttribute(DB, cell, "office:value-type")
						celldatevalue := xmlDB.GetNodeAttribute(DB, cell, "office:date-value")
						cellvalue := xmlDB.GetNodeAttribute(DB, cell, "office:value")
						cellformula := xmlDB.GetNodeAttribute(DB, cell, "table:formula")

						Cell_paras, _ := xmlDB.GetNode(DB, cell, "text:p")
						celltext := ""
						for _, Cell_para := range Cell_paras {
							childnodes := xmlDB.ChildNodes(DB, Cell_para)
							if len(childnodes) > 0 {
								for _, child := range childnodes {
									nodeName := xmlDB.GetNodeName(DB, child)
									if nodeName == "text:s" {
										cell_space := 1
										cell_spaceTXT := xmlDB.GetNodeAttribute(DB, child, "text:c")
										if len(strings.TrimSpace(cell_spaceTXT)) > 0 {
											value, err := strconv.Atoi(cell_spaceTXT)
											if err != nil {
												return odsfileContents, err
											}
											cell_space = value
										}
										for {
											if cell_space == 0 {
												break
											}
											celltext = celltext + " "
											cell_space--
										}
									}
									celltext = celltext + xmlDB.GetNodeValue(DB, child)
								}
							} else {
								if len(celltext) > 0 {
									celltext = celltext + "\n" + xmlDB.GetNodeValue(DB, Cell_para)
								} else {
									celltext = xmlDB.GetNodeValue(DB, Cell_para)
								}

							}
						}

						celltext = ReplaceHTMLSpecialEntities(celltext)
						cellvalue = ReplaceHTMLSpecialEntities(cellvalue)
						if len(celltext) == 0 {
							blankcells = blankcells + cell_repetition
						} else {
							//insert blankcells before newcells
							for {
								if blankcells == 0 {
									break
								}
								csvcells = append(csvcells, Cell{"", "", "", "", ""})
								blankcells--

							}
							for {
								if cell_repetition == 0 {
									break
								}
								csvcells = append(csvcells, Cell{celltype, cellvalue, celldatevalue, cellformula, celltext})
								cell_repetition--
								rowisblank = false
							}
						}

					}
					if rowisblank {
						blankrows = blankrows + row_repetition
					} else {
						//insert blankrows before newrows
						for {
							if blankrows == 0 {
								break
							}
							csvrows = append(csvrows, Row{[]Cell{}})
							blankrows--
						}
						for {
							if row_repetition == 0 {
								break
							}
							csvrows = append(csvrows, Row{csvcells})
							row_repetition--
						}

					}

				}
				csvSpreadSheets = append(csvSpreadSheets, Sheet{tablename, csvrows})
			}
			odsfileContents.Sheets = csvSpreadSheets
			return odsfileContents, err
		}
	}

	return odsfileContents, err
}
