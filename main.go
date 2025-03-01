package main

import (
	"encoding/json"
	"flag"
	"fmt"
	"github.com/xuri/excelize/v2"
	"log"
	"net/http"
	"os"
	"sort"
	"strconv"
	"strings"
	"sync"
	"time"
)

type StudentRecord struct {
	Emplid         string
	Name           string
	Branch         string
	Batch          string
	ClassNo        string
	Quiz           float64
	MidSem         float64
	LabTest        float64
	WeeklyLabs     float64
	PreCompre      float64
	Compre         float64
	TotalGiven     float64
	TotalComputed  float64
	HasDiscrepancy bool
}

type ComponentRank struct {
	Emplid string
	Name   string
	Marks  float64
	Rank   int
}

type SummaryReport struct {
	GeneralAverages  map[string]float64         `json:"generalAverages"`
	BranchAverages   map[string]float64         `json:"branchAverages"`
	ComponentToppers map[string][]ComponentRank `json:"componentToppers"`
	Discrepancies    []StudentRecord            `json:"discrepancies,omitempty"`
}

const (
	EmplIdCol     = 0
	NameCol       = 1
	BranchCol     = 2
	BatchCol      = 3
	ClassNoCol    = 4
	QuizCol       = 5
	MidSemCol     = 6
	LabTestCol    = 7
	WeeklyLabsCol = 8
	PreCompreCol  = 9
	CompreCol     = 10
	TotalCol      = 11
)

var componentInfo = map[string]float64{
	"Quiz":       30,
	"MidSem":     75,
	"LabTest":    60,
	"WeeklyLabs": 30,
	"PreCompre":  195,
	"Compre":     105,
	"Total":      300,
}

func main() {
	exportFlag := flag.String("export", "", "Export format (e.g., json)")
	classFilter := flag.String("class", "", "Filter by class number")
	flag.Parse()

	args := flag.Args()
	if len(args) < 1 {
		log.Fatal("Error: Excel file path or URL not provided. Usage: go run main.go [--export=json] [--class=XXXX] path/to/file.xlsx OR http://example.com/file.xlsx")
	}
	filePathOrURL := args[0]

	var records []StudentRecord
	var err error

	if strings.HasPrefix(filePathOrURL, "http://") || strings.HasPrefix(filePathOrURL, "https://") {
		fmt.Println("Processing file from URL:", filePathOrURL)
		records, err = processExcelFromURL(filePathOrURL, *classFilter)
		if err != nil {
			log.Fatalf("Error processing file from URL: %v", err)
		}
	} else {
		if _, err := os.Stat(filePathOrURL); os.IsNotExist(err) {
			log.Fatalf("Error: File not found at %s", filePathOrURL)
		}
		records, err = processExcelFile(filePathOrURL, *classFilter)
		if err != nil {
			log.Fatalf("Error processing file: %v", err)
		}
	}

	report := generateReport(records)

	printReport(report, records)

	if *exportFlag == "json" {
		exportToJSON(report)
	}
}

func processExcelFromURL(url string, classFilter string) ([]StudentRecord, error) {
	if strings.Contains(url, "docs.google.com/spreadsheets") {
		parts := strings.Split(url, "/")
		var sheetID string
		for i, part := range parts {
			if part == "d" && i+1 < len(parts) {
				sheetID = parts[i+1]
				break
			}
		}
		
		if idx := strings.Index(sheetID, "?"); idx != -1 {
			sheetID = sheetID[:idx]
		}
		if idx := strings.Index(sheetID, "#"); idx != -1 {
			sheetID = sheetID[:idx]
		}
		
		url = fmt.Sprintf("https://docs.google.com/spreadsheets/d/%s/export?format=xlsx", sheetID)
	}

	client := &http.Client{
		Timeout: 30 * time.Second,
	}

	resp, err := client.Get(url)
	if err != nil {
		return nil, fmt.Errorf("failed to fetch URL: %v", err)
	}
	defer resp.Body.Close()

	if resp.StatusCode != http.StatusOK {
		return nil, fmt.Errorf("bad status: %s", resp.Status)
	}

	xlsx, err := excelize.OpenReader(resp.Body)
	if err != nil {
		return nil, fmt.Errorf("failed to open Excel file from URL: %v", err)
	}
	defer xlsx.Close()

	return processExcelData(xlsx, classFilter)
}

func processExcelFile(filePath, classFilter string) ([]StudentRecord, error) {
	xlsx, err := excelize.OpenFile(filePath)
	if err != nil {
		return nil, fmt.Errorf("failed to open file: %v", err)
	}
	defer xlsx.Close()

	return processExcelData(xlsx, classFilter)
}

func processExcelData(xlsx *excelize.File, classFilter string) ([]StudentRecord, error) {
	sheets := xlsx.GetSheetList()
	if len(sheets) == 0 {
		return nil, fmt.Errorf("no sheets found in the Excel file")
	}
	sheetName := sheets[0]
	rows, err := xlsx.GetRows(sheetName)
	if err != nil {
		return nil, fmt.Errorf("failed to get rows from sheet: %v", err)
	}

	var records []StudentRecord
	startRow := 0

	for i, row := range rows {
		if len(row) > 0 && isNumeric(row[EmplIdCol]) {
			startRow = i
			break
		}
	}

	for i := startRow; i < len(rows); i++ {
		row := rows[i]

		if len(row) <= TotalCol || row[EmplIdCol] == "" {
			continue
		}

		record, err := parseStudentRecord(row)
		if err != nil {
			log.Printf("Warning: Skipping row %d - %v", i+1, err)
			continue
		}

		if classFilter != "" && record.ClassNo != classFilter {
			continue
		}

		records = append(records, record)
	}

	if len(records) == 0 {
		return nil, fmt.Errorf("no valid records found in the file")
	}

	return records, nil
}

func parseStudentRecord(row []string) (StudentRecord, error) {
	record := StudentRecord{
		Emplid:  row[EmplIdCol],
		Name:    row[NameCol],
		Branch:  row[BranchCol],
		Batch:   row[BatchCol],
		ClassNo: row[ClassNoCol],
	}

	var err error

	record.Quiz, err = parseFloat(row[QuizCol])
	if err != nil {
		return record, fmt.Errorf("invalid Quiz mark: %v", err)
	}

	record.MidSem, err = parseFloat(row[MidSemCol])
	if err != nil {
		return record, fmt.Errorf("invalid MidSem mark: %v", err)
	}

	record.LabTest, err = parseFloat(row[LabTestCol])
	if err != nil {
		return record, fmt.Errorf("invalid LabTest mark: %v", err)
	}

	record.WeeklyLabs, err = parseFloat(row[WeeklyLabsCol])
	if err != nil {
		return record, fmt.Errorf("invalid WeeklyLabs mark: %v", err)
	}

	record.PreCompre, err = parseFloat(row[PreCompreCol])
	if err != nil {
		return record, fmt.Errorf("invalid PreCompre mark: %v", err)
	}

	record.Compre, err = parseFloat(row[CompreCol])
	if err != nil {
		return record, fmt.Errorf("invalid Compre mark: %v", err)
	}

	record.TotalGiven, err = parseFloat(row[TotalCol])
	if err != nil {
		return record, fmt.Errorf("invalid Total mark: %v", err)
	}

	record.TotalComputed = record.Quiz + record.MidSem + record.LabTest +
		record.WeeklyLabs + record.PreCompre + record.Compre

	const epsilon = 0.01
	if abs(record.TotalGiven-record.TotalComputed) > epsilon {
		record.HasDiscrepancy = true
	}

	return record, nil
}

func generateReport(records []StudentRecord) SummaryReport {
	var report SummaryReport
	var wg sync.WaitGroup
	var mu sync.Mutex

	report.GeneralAverages = make(map[string]float64)
	report.BranchAverages = make(map[string]float64)
	report.ComponentToppers = make(map[string][]ComponentRank)

	wg.Add(1)
	go func() {
		defer wg.Done()

		componentSums := map[string]float64{
			"Quiz": 0, "MidSem": 0, "LabTest": 0, "WeeklyLabs": 0,
			"PreCompre": 0, "Compre": 0, "Total": 0,
		}

		for _, record := range records {
			componentSums["Quiz"] += record.Quiz
			componentSums["MidSem"] += record.MidSem
			componentSums["LabTest"] += record.LabTest
			componentSums["WeeklyLabs"] += record.WeeklyLabs
			componentSums["PreCompre"] += record.PreCompre
			componentSums["Compre"] += record.Compre
			componentSums["Total"] += record.TotalGiven
		}

		count := float64(len(records))
		mu.Lock()
		for component, sum := range componentSums {
			report.GeneralAverages[component] = sum / count
		}
		mu.Unlock()
	}()

	wg.Add(1)
	go func() {
		defer wg.Done()

		branchSums := make(map[string]float64)
		branchCounts := make(map[string]int)

		for _, record := range records {
			if strings.Contains(record.Batch, "2024") {
				if !strings.Contains(record.Branch, "&") && !strings.Contains(record.Branch, "+") {
					branchSums[record.Branch] += record.TotalGiven
					branchCounts[record.Branch]++
				}
			}
		}

		mu.Lock()
		for branch, sum := range branchSums {
			if branchCounts[branch] > 0 {
				report.BranchAverages[branch] = sum / float64(branchCounts[branch])
			}
		}
		mu.Unlock()
	}()

	components := []string{"Quiz", "MidSem", "LabTest", "WeeklyLabs", "PreCompre", "Compre", "Total"}
	for _, component := range components {
		wg.Add(1)
		go func(comp string) {
			defer wg.Done()

			var rankings []ComponentRank

			for _, record := range records {
				var marks float64
				switch comp {
				case "Quiz":
					marks = record.Quiz
				case "MidSem":
					marks = record.MidSem
				case "LabTest":
					marks = record.LabTest
				case "WeeklyLabs":
					marks = record.WeeklyLabs
				case "PreCompre":
					marks = record.PreCompre
				case "Compre":
					marks = record.Compre
				case "Total":
					marks = record.TotalGiven
				}

				rankings = append(rankings, ComponentRank{
					Emplid: record.Emplid,
					Name:   record.Name,
					Marks:  marks,
				})
			}

			sort.Slice(rankings, func(i, j int) bool {
				return rankings[i].Marks > rankings[j].Marks
			})

			top := min(3, len(rankings))
			for i := 0; i < top; i++ {
				rankings[i].Rank = i + 1
			}

			mu.Lock()
			report.ComponentToppers[comp] = rankings[:top]
			mu.Unlock()
		}(component)
	}

	wg.Add(1)
	go func() {
		defer wg.Done()

		var discrepancies []StudentRecord
		for _, record := range records {
			if record.HasDiscrepancy {
				discrepancies = append(discrepancies, record)
			}
		}

		mu.Lock()
		report.Discrepancies = discrepancies
		mu.Unlock()
	}()

	wg.Wait()

	return report
}

func printReport(report SummaryReport, records []StudentRecord) {
	fmt.Println("========== GRADE ANALYSIS REPORT ==========")
	fmt.Printf("Total Records Processed: %d\n\n", len(records))

	fmt.Println("=== DISCREPANCIES ===")
	if len(report.Discrepancies) == 0 {
		fmt.Println("No discrepancies found.")
	} else {
		fmt.Printf("Found %d discrepancies:\n", len(report.Discrepancies))
		for i, record := range report.Discrepancies {
			fmt.Printf("%d. Emplid: %s, Name: %s\n", i+1, record.Emplid, record.Name)
			fmt.Printf("   Given Total: %.2f, Computed Total: %.2f, Difference: %.2f\n",
				record.TotalGiven, record.TotalComputed, record.TotalGiven-record.TotalComputed)
		}
	}
	fmt.Println()

	fmt.Println("=== GENERAL AVERAGES ===")
	for component, average := range report.GeneralAverages {
		maxMark := componentInfo[component]
		percentage := (average / maxMark) * 100
		fmt.Printf("%s: %.2f / %.0f (%.2f%%)\n", component, average, maxMark, percentage)
	}
	fmt.Println()

	fmt.Println("=== BRANCH-WISE AVERAGES (2024 Single Degree) ===")
	if len(report.BranchAverages) == 0 {
		fmt.Println("No branch data available.")
	} else {
		var branches []string
		for branch := range report.BranchAverages {
			branches = append(branches, branch)
		}
		sort.Strings(branches)

		for _, branch := range branches {
			average := report.BranchAverages[branch]
			fmt.Printf("%s: %.2f / 300 (%.2f%%)\n", branch, average, (average/300)*100)
		}
	}
	fmt.Println()

	fmt.Println("=== TOP 3 STUDENTS BY COMPONENT ===")
	components := []string{"Quiz", "MidSem", "LabTest", "WeeklyLabs", "PreCompre", "Compre", "Total"}
	for _, component := range components {
		fmt.Printf("--- %s ---\n", component)
		toppers := report.ComponentToppers[component]

		if len(toppers) == 0 {
			fmt.Println("No data available.")
			continue
		}

		maxMark := componentInfo[component]
		for _, topper := range toppers {
			rankText := ""
			switch topper.Rank {
			case 1:
				rankText = "1st"
			case 2:
				rankText = "2nd"
			case 3:
				rankText = "3rd"
			default:
				rankText = fmt.Sprintf("%dth", topper.Rank)
			}

			fmt.Printf("%s: %s (%s) - %.2f / %.0f (%.2f%%)\n",
				rankText, topper.Name, topper.Emplid,
				topper.Marks, maxMark, (topper.Marks/maxMark)*100)
		}
		fmt.Println()
	}
}

func exportToJSON(report SummaryReport) {
	jsonData, err := json.MarshalIndent(report, "", "  ")
	if err != nil {
		log.Printf("Error creating JSON: %v", err)
		return
	}

	err = os.WriteFile("grade_report.json", jsonData, 0644)
	if err != nil {
		log.Printf("Error writing JSON file: %v", err)
		return
	}

	fmt.Println("Report exported to grade_report.json")
}

func parseFloat(s string) (float64, error) {
	s = strings.TrimSpace(s)
	if s == "" {
		return 0, nil
	}
	return strconv.ParseFloat(s, 64)
}

func isNumeric(s string) bool {
	_, err := strconv.ParseFloat(strings.TrimSpace(s), 64)
	return err == nil
}

func abs(x float64) float64 {
	if x < 0 {
		return -x
	}
	return x
}

func min(a, b int) int {
	if a < b {
		return a
	}
	return b
}
