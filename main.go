package main

import (
	"fmt"
	"io"
	"net/http"
	"os"
	"os/exec"
	"path/filepath"
	"time"
)

const tempDir = "./tmp" // Directory for temporary files

func main() {
	// Ensure the temporary directory exists
	if err := os.MkdirAll(tempDir, os.ModePerm); err != nil {
		fmt.Println("Failed to create temp directory:", err)
		return
	}

	// Start the file cleanup goroutine
	go cleanupOldFiles(tempDir, 1*time.Hour)

	http.HandleFunc("/", handleHealthCheck)
	http.HandleFunc("/convert", handleConvert)

	fmt.Println("Starting server on :5000")
	if err := http.ListenAndServe(":5000", nil); err != nil {
		fmt.Println("Failed to start server:", err)
	}
}

func handleHealthCheck(w http.ResponseWriter, r *http.Request) {
	w.Write([]byte("OK"))
}

func handleConvert(w http.ResponseWriter, r *http.Request) {
	// Ensure the request method is POST
	if r.Method != http.MethodPost {
		http.Error(w, "Only POST method is allowed", http.StatusMethodNotAllowed)
		return
	}
	// Ensure it has proper token
    expectedToken := os.Getenv("API_SECRET_TOKEN")
    if expectedToken != "" && r.Header.Get("X-API-Key") != expectedToken {
        http.Error(w, "Unauthorized", http.StatusUnauthorized)
        return
    }

	// Parse the uploaded file
	file, _, err := r.FormFile("file")
	if err != nil {
		http.Error(w, "Failed to read uploaded file", http.StatusBadRequest)
		return
	}
	defer file.Close()

	// Save the Excel file to a temporary location
	baseName := time.Now().Format("20060102150405") // Timestamp format
	inputFilePath := filepath.Join(tempDir, baseName+".xlsx")
	outputFilePath := filepath.Join(tempDir, baseName+".pdf")

	inputFile, err := os.Create(inputFilePath)
	if err != nil {
		http.Error(w, "Failed to create temporary file", http.StatusInternalServerError)
		return
	}
	defer inputFile.Close()
	defer os.Remove(inputFilePath)

	_, err = io.Copy(inputFile, file)
	if err != nil {
		http.Error(w, "Failed to save uploaded file", http.StatusInternalServerError)
		return
	}

	// Normalize spacing in docx before converting
	normalizeCmd := exec.Command("python3", "/app/normalize.py", inputFilePath)
	output, err := normalizeCmd.CombinedOutput()
	fmt.Println("Normalize output:", string(output))
	if err != nil {
		fmt.Println("Normalize error:", err)
	}

	// Convert the Excel file to PDF using LibreOffice
	cmd := exec.Command("soffice", "--headless", "--convert-to", "pdf:writer_pdf_Export", inputFilePath, "--outdir", tempDir)
	if err := cmd.Run(); err != nil {
		fmt.Println(err)
		http.Error(w, "Failed to convert file to PDF", http.StatusInternalServerError)
		return
	}
	defer os.Remove(outputFilePath)

	// Read the converted PDF file
	pdfFile, err := os.Open(outputFilePath)
	if err != nil {
		fmt.Println(err)
		http.Error(w, "Failed to read converted PDF", http.StatusInternalServerError)
		return
	}
	defer pdfFile.Close()

	// Write the PDF file as response
	w.Header().Set("Content-Type", "application/pdf")
	w.Header().Set("Content-Disposition", `attachment; filename="output.pdf"`)

	if _, err := io.Copy(w, pdfFile); err != nil {
		http.Error(w, "Failed to write PDF to response", http.StatusInternalServerError)
		return
	}
}

// cleanupOldFiles removes files older than the specified duration from the given directory
func cleanupOldFiles(dir string, maxAge time.Duration) {
	for {
		time.Sleep(1 * time.Hour) // Check every minute

		files, err := os.ReadDir(dir)
		if err != nil {
			fmt.Println("Failed to read temp directory:", err)
			continue
		}

		for _, file := range files {
			filePath := filepath.Join(dir, file.Name())
			info, err := os.Stat(filePath)
			if err != nil {
				fmt.Println("Failed to get file info:", err)
				continue
			}

			// Check if the file is older than maxAge
			if time.Since(info.ModTime()) > maxAge {
				if err := os.Remove(filePath); err != nil {
					fmt.Println("Failed to delete file:", err)
				} else {
					fmt.Println("Deleted old file:", filePath)
				}
			}
		}
	}
}
