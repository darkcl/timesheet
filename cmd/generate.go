// Copyright © 2018 NAME HERE <EMAIL ADDRESS>
//
// Licensed under the Apache License, Version 2.0 (the "License");
// you may not use this file except in compliance with the License.
// You may obtain a copy of the License at
//
//     http://www.apache.org/licenses/LICENSE-2.0
//
// Unless required by applicable law or agreed to in writing, software
// distributed under the License is distributed on an "AS IS" BASIS,
// WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
// See the License for the specific language governing permissions and
// limitations under the License.

package cmd

import (
	"fmt"
	"time"

	"github.com/rickar/cal"
	"github.com/spf13/cobra"
	"github.com/tealeg/xlsx"
)

// generateCmd represents the generate command
var generateCmd = &cobra.Command{
	Use:   "generate",
	Short: "Generate Timesheet",
	Run: func(cmd *cobra.Command, args []string) {
		c := cal.NewCalendar()

		c.AddHoliday(
			cal.NewHoliday(time.January, 1),
			cal.NewHoliday(time.February, 16),
			cal.NewHoliday(time.February, 17),
			cal.NewHoliday(time.February, 18),
			cal.NewHoliday(time.February, 19),
			cal.NewHoliday(time.March, 30),
			cal.NewHoliday(time.March, 31),
			cal.NewHoliday(time.April, 2),
			cal.NewHoliday(time.April, 5),
			cal.NewHoliday(time.May, 1),
			cal.NewHoliday(time.May, 22),
			cal.NewHoliday(time.June, 18),
			cal.NewHoliday(time.July, 1),
			cal.NewHoliday(time.July, 2),
			cal.NewHoliday(time.September, 25),
			cal.NewHoliday(time.October, 1),
			cal.NewHoliday(time.October, 17),
			cal.NewHoliday(time.December, 25),
			cal.NewHoliday(time.December, 26),
		)

		var file *xlsx.File
		var sheet *xlsx.Sheet
		var row *xlsx.Row
		var cell *xlsx.Cell
		var err error

		file = xlsx.NewFile()
		sheet, err = file.AddSheet("Sheet1")
		if err != nil {
			fmt.Printf(err.Error())
		}
		currentDate := time.Now().Local()

		start := time.Date(currentDate.Year(), currentDate.Month(), 1, 0, 0, 0, 0, currentDate.Location())

		// set d to starting date and keep adding 1 day to it as long as month doesn't change
		for d := start; d.Month() == start.Month(); d = d.AddDate(0, 0, 1) {
			row = sheet.AddRow()
			cell = row.AddCell()
			if c.IsWorkday(d) {
				cell.SetInt(8)
			} else {
				cell.Value = ""
			}

			err = file.Save("day.xlsx")
			if err != nil {
				fmt.Printf(err.Error())
			}
		}
	},
}

func init() {
	rootCmd.AddCommand(generateCmd)

	// Here you will define your flags and configuration settings.

	// Cobra supports Persistent Flags which will work for this command
	// and all subcommands, e.g.:
	// generateCmd.PersistentFlags().String("foo", "", "A help for foo")

	// Cobra supports local flags which will only run when this command
	// is called directly, e.g.:
	// generateCmd.Flags().BoolP("toggle", "t", false, "Help message for toggle")
}
