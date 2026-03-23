# AGENTS.md

This file provides guidance to Codex (Codex.ai/code) when working with code in this repository.

## Common Commands

### Building and Testing
```bash
# Restore dependencies
dotnet restore

# Build the project
dotnet build

# Run tests
dotnet test

# Run the application (uses next quarter by default)
dotnet run

# Run the application for a specific date
dotnet run 2/1/2025
```

### Testing Individual Components
```bash
# Run specific test file
dotnet test tests/CreateSoundRoomSchedule.Tests/QuarterTests.cs

# Run tests with verbose output
dotnet test --verbosity normal
```

## Architecture Overview

This is a .NET 10 console application that generates Excel calendars for church A/V scheduling. The application:

1. **Fetches service data** from Planning Center API using Planning Center personal access token credentials
2. **Generates quarterly calendars** spanning 3 months with A/V team assignments
3. **Includes holiday information** from external API with custom church holidays
4. **Outputs Excel files** with print-optimized formatting to the desktop

### Key Components

- **Program.cs**: Main entry point and orchestration only
- **ExcelScheduleBuilder.cs**: Excel workbook and worksheet generation logic
- **PlanningCenter.cs**: API client for Planning Center services and team data with defensive JSON parsing
- **QuarterCalculator.cs**: Quarter calculation utilities (partial class of Program)
- **Holiday.cs**: Public holiday API integration with church-specific additions and graceful fallback on API failure
- **Constants.cs**: Excel formatting constants and configuration values

### Data Flow
1. Parse the optional date argument; missing or invalid input defaults to the next quarter
2. Fetch services from Planning Center API for the quarter date range
3. Retrieve team assignments for each service (sound, stream, slides roles)
4. Fetch public holidays and add church-specific dates
5. Generate Excel workbook with one worksheet per month
6. Apply print settings optimized for landscape printing and fit-to-width printing

## Configuration

### User Secrets
The application requires Planning Center API credentials stored as user secrets:
- `PlanningCenter:clientId` - Planning Center PAT client ID
- `PlanningCenter:clientSecret` - Corresponding Planning Center PAT secret

### External Dependencies
- **EPPlus**: Excel file generation (NonCommercial license)
- **Planning Center API**: Service and team data
- **date.nager.at**: Public holiday data

## Testing

Tests use xUnit framework and focus on:
- Quarter calculation logic
- Planning Center response parsing
- Excel cell formatting and print setting validation

The test project references the main project to test internal methods.
