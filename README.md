# ExcelWriter v2

A powerful Go package for exporting structured data to Excel files with rich formatting options, data validation, and customizable styling.

## Features

- Export slices of structs to Excel files
- Rich cell styling (font, color, alignment, etc.)
- Column-level configuration (width, visibility, freezing)
- Data validation rules
- Sheet protection with password
- Auto-filtering support
- Header row freezing
- Custom number and date formatting
- Timezone-aware timestamp handling
- Decimal precision control
- Custom value formatting via Export interface
- Boolean value mapping

## Installation

```bash
go get github.com/shopspring/decimal
go get github.com/xuri/excelize/v2
```

## Usage

```go
// Custom type implementing Export interface
type Status string

func (s Status) ExcelString() string {
    return string(s)
}

func (s Status) ExcelStyle() *CellStyle {
    switch s {
    case "Active":
        return &CellStyle{
            FontColor: "#008000",  // Green
            Bold:      true,
        }
    case "Inactive":
        return &CellStyle{
            FontColor: "#FF0000",  // Red
            Italic:    true,
        }
    default:
        return nil
    }
}

type Product struct {
    ID          int             `excel:"title=Product ID;width=15;style={'bold':true,'alignment':'center'}"`
    Name        string          `excel:"title=Product Name;width=30;validation={'type':'textLength','operator':'between','formula1':'3','formula2':'50'}"`
    Price       decimal.Decimal `excel:"title=Price;format=#,##0.00;style={'numberFormat':'$#,##0.00'}"`
    Category    string          `excel:"title=Category;validation={'type':'list','allowList':['Electronics','Books','Clothing']}"`
    InStock     bool           `excel:"title=库存状态;width=10;valuemap={'true':'有货','false':'断货'}"`
    LastUpdated time.Time      `excel:"title=Last Updated;timestamp;format=2006-01-02 15:04"`
    Status      Status         `excel:"title=Status;width=15"` // Uses Export interface
    Rating      float64        `excel:"title=Rating;validation={'type':'decimal','operator':'between','formula1':'0','formula2':'5'};format=#0.0"`
    Internal    string         `excel:"omit"`
}

// Create exporter with advanced options
excel := excelwriter.New(
    excelwriter.WithTimezone("Asia/Hong_Kong"),
    excelwriter.WithSheetName("Products"),
    excelwriter.WithDatetimeFormat("2006-01-02 15:04:05"),
    excelwriter.WithDecimalDigits(2),
    excelwriter.WithHeaderStyle(&excelwriter.CellStyle{
        Bold:       true,
        FontSize:   12,
        Background: "#CCCCCC",
        Alignment:  "center",
    }),
    excelwriter.WithAutoFilter(true),
    excelwriter.WithFreezeHeader(true),
    excelwriter.WithPassword("secret"),
)

products := []Product{
    {
        ID:          1,
        Name:        "Laptop Pro",
        Price:       decimal.NewFromFloat(999.99),
        Category:    "Electronics",
        InStock:     true,
        LastUpdated: time.Now(),
        Status:      Status("Active"),
        Rating:      4.5,
    },
    // ... more products
}

// Export data
reader, err := excel.Export(products)
if err != nil {
    log.Fatal(err)
}
```

## Custom Value Formatting (Export Interface)

The package provides an `Export` interface that allows types to control their Excel representation and styling:

```go
type Export interface {
    ExcelString() string              // Returns the string representation for Excel
    ExcelStyle() *CellStyle          // Returns custom styling for the cell (optional)
}
```

Example implementation:
```go
type Priority int

func (p Priority) ExcelString() string {
    switch p {
    case 1:
        return "Low"
    case 2:
        return "Medium"
    case 3:
        return "High"
    default:
        return "Unknown"
    }
}

func (p Priority) ExcelStyle() *CellStyle {
    switch p {
    case 1:
        return &CellStyle{Background: "#90EE90"}  // Light green
    case 2:
        return &CellStyle{Background: "#FFD700"}  // Gold
    case 3:
        return &CellStyle{Background: "#FFB6C1", Bold: true}  // Light red
    default:
        return nil
    }
}
```

## Struct Tags

The package uses a semicolon-separated tag format with rich options:

```go
`excel:"title=Column Title;width=20;format=#,##0.00;style={'bold':true,'fontColor':'red'};validation={'type':'decimal','operator':'between','formula1':'0','formula2':'1000'};freeze;hidden"`
```

Supported tag options:

- `title=value`: Set column header name
- `width=float`: Set column width
- `format=string`: Set cell format pattern
- `style=json`: Set cell style (JSON object)
- `validation=json`: Set data validation rules (JSON object)
- `timestamp`: Format as datetime
- `omit`: Exclude from output
- `freeze`: Freeze this column
- `hidden`: Hide this column
- `valuemap=json`: Map values to display strings (useful for booleans)

## Style Options

```go
type CellStyle struct {
    FontSize     float64 // Font size in points
    FontColor    string  // HTML color code
    Background   string  // HTML color code
    Bold         bool    // Bold text
    Italic       bool    // Italic text
    Alignment    string  // left, center, right
    NumberFormat string  // Custom number format
}
```

## Data Validation

```go
type DataValidation struct {
    Type      string   // decimal, whole, date, time, textLength, list
    Operator  string   // between, notBetween, equal, notEqual, greaterThan, lessThan
    Formula1  string   // First value or formula
    Formula2  string   // Second value or formula (for between operations)
    AllowList []string // Values for dropdown lists
}
```

## Configuration Options

- `WithTimezone(timezone string)`: Set timezone for timestamps
- `WithSheetName(name string)`: Set worksheet name
- `WithDecimalDigits(digits int32)`: Set decimal precision
- `WithDatetimeFormat(format string)`: Set datetime format
- `WithDefaultStyle(*CellStyle)`: Set default cell style
- `WithHeaderStyle(*CellStyle)`: Set header row style
- `WithPassword(password string)`: Set worksheet protection
- `WithAutoFilter(bool)`: Enable auto-filtering
- `WithFreezeHeader(bool)`: Freeze header row
- `WithActiveSheet(bool)`: Make sheet active when opened
