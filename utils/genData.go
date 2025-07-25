package utils

import (
	"encoding/json"
	"fmt" // Added for debugging print statements
	"os"
	"strings"

	"github.com/xuri/excelize/v2"
)

// GenerateJson reads the timetable.xlsx file, extracts sheet names and class/subgroup names,
// and then converts this data into a JSON format stored in data.json.
func GenerateJson() {
	// Debugging: Confirm that the function is attempting to open the Excel file.
	fmt.Println("Attempting to open timetable.xlsx...")

	// Open the Excel file.
	f, err := excelize.OpenFile("timetable.xlsx")
	// Ensure the file is closed when the function exits.
	defer func() {
		if err = f.Close(); err != nil {
			// If there's an error closing the file, panic (crash) the program.
			// This indicates a serious issue that needs immediate attention.
			panic(err)
		}
	}()
	// Handle any error that occurred during file opening.
	HandleError(err)

	// Get a list of all sheet names in the Excel file.
	sheets := f.GetSheetList()
	// Initialize a map to store classes for each sheet.
	// The structure is: sheetName -> (column_index -> subgroup_name)
	classes := make(map[string]map[int]string)

	// Iterate over each sheet found in the Excel file.
	for _, sheetName := range sheets {
		// Initialize a temporary map to store column_index -> subgroup_name for the current sheet.
		temp := make(map[int]string)
		// Get all rows from the current sheet.
		rows, err := f.GetRows(sheetName)
		// Handle any error that occurred while getting rows.
		HandleError(err)

		// Check if row 5 (index 4) exists, as this is where subgroups are located.
		if len(rows) > 4 {
			// Iterate through each cell in row 5.
			for colIndex, cellValue := range rows[4] { // row 5 is at index 4
				// Trim whitespace from the cell value.
				trimmedValue := strings.TrimSpace(cellValue)

				// Filter out common non-class headers that might appear in row 5.
				// This ensures only actual class/subgroup codes are captured.
				if trimmedValue != "" &&
					trimmedValue != "DAY" &&
					trimmedValue != "HOURS" &&
					trimmedValue != "SR NO" &&
					trimmedValue != "SR.NO" &&
					trimmedValue != "TUTORIAL" &&
					trimmedValue != "LECTURE" &&
					trimmedValue != "PRACTICAL" &&
					trimmedValue != "BRANCH" { // Added "BRANCH" as it appears in some headers
					// Store the subgroup name with its 1-based column index.
					temp[colIndex+1] = trimmedValue
				}
			}
		}
		// Assign the collected subgroups for the current sheet.
		classes[sheetName] = temp
	}

	// Convert the extracted Excel data into JSON format.
	ExcelToJson(classes, f)
}

// ExcelToJson takes the extracted class data and the Excel file object,
// then processes it to create the final data.json file.
func ExcelToJson(classes map[string]map[int]string, f *excelize.File) {
	// Open or create the data.json file. os.Create will create the file if it doesn't exist,
	// or truncate (clear) it if it does. This is safer for ensuring the file is writable.
	file, err := os.Create("./data.json") // Changed from os.OpenFile to os.Create
	HandleError(err)
	// Ensure the JSON file is closed when the function exits.
	defer file.Close()

	// Initialize the main data structure for the JSON output.
	// This will hold: sheetName -> (subgroup_name -> timetable_data)
	data := make(map[string]map[string][][]Data)

	// Iterate over each sheet and its collected subgroups.
	for sheetName, subgroupsInSheet := range classes { // sheetName is like "2ND YEAR B", subgroupsInSheet is map[int]string (colIndex -> subgroup_name)
		// Initialize a temporary map for the current sheet's data.
		tempSheetData := make(map[string][][]Data)
		// Iterate over each subgroup (column index and subgroup name).
		for colIndex, subgroupName := range subgroupsInSheet {
			// Get the actual timetable data for this subgroup from the Excel file.
			// This calls the GetTableData function from data.go.
			timetableData := GetTableData(sheetName, colIndex, f)
			// Assign the retrieved timetable data to the subgroup name.
			tempSheetData[strings.Trim(subgroupName, " ")] = timetableData
		}
		// Assign the current sheet's data to the main data map.
		data[strings.Trim(sheetName, " ")] = tempSheetData
	}

	// Marshal (convert) the Go data structure into a pretty-printed JSON byte array.
	dj, _ := json.MarshalIndent(data, "", " ")
	// Write the JSON data to the data.json file.
	_, err = file.Write(dj)
	// Handle any error that occurred during writing.
	HandleError(err)
}
