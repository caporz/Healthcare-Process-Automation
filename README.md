# VBA Macros | Administrative Process Automation

This repository collects automation projects developed with **Excel VBA**, with the support of Excel and Power Query where appropriate.

The projects address practical administrative, accounting, reporting, and management-control activities.

The underlying idea is simple: repetitive and time-consuming tasks can often be redesigned and automated, allowing professionals to focus on analysis, control, and decision-making.

## Projects

### 1. ALPI – Reporting Automation

Automation of reporting activities related to intramoenia and professional services.

The project includes data processing and aggregation through VBA.

**Main focus:**
- Processing structured data.
- Automating reporting activities.
- Supporting the preparation of summary outputs.

**Relevant limitation:** some calculations depend on the ordering of invoices. Input data must therefore satisfy the assumptions documented in the project.

### 2. Risk Classification – Company Clustering

A VBA project focused on the classification and clustering of companies according to risk-related criteria.

The project explores how structured data processing can support risk assessment activities.

**Main focus:**
- Processing company-level information.
- Applying classification criteria.
- Supporting the identification and grouping of risk profiles.

The resulting classification should be interpreted in light of the criteria and assumptions used. It is not a substitute for professional risk assessment.

### 3. Extramoenia – Administrative Workflow Automation

An automation workflow designed to support administrative activities related to extramoenia services.

The project combines data extraction and processing with Excel, Power Query, Pivot Tables, VBA, and Outlook.

**Main workflow:**
- Processing source data.
- Organizing information by professional.
- Preparing individual summaries.
- Supporting the generation of communications through Outlook.

**Relevant limitation:** the workflow may not generate a communication for professionals whose activity consists exclusively of outpatient services. The input data and business rules must be checked accordingly.

### 4. Project Monitoring – Budget & Reporting

A VBA project designed to support project monitoring and economic reporting.

**Main focus:**
- Organizing project-related information.
- Monitoring budgets and economic data.
- Supporting the identification of deviations.
- Automating recurring reporting activities.

## Project Structure

```text
VBA-Macros/
├── ALPI/
│   ├── README.md
│   └── RiepilogoALPI.bas
├── Clusterizzazione-rischio/
│   ├── README.md
│   └── RischioSocieta.bas
├── Extramoenia/
│   ├── README.md
│   └── Extramoenia.bas
└── Monitoraggio-progetti/
    ├── README.md
    └── MonitoraggioProgetti.bas
```

## Design Principles

The projects follow a practical approach to automation:

1. Understand the existing process.
2. Identify repetitive activities and potential sources of error.
3. Define the required inputs and outputs.
4. Implement the automation.
5. Preserve appropriate checks and exception handling.
6. Evaluate the operational benefits.

## Technologies

- Microsoft Excel
- VBA
- Power Query
- Pivot Tables
- Microsoft Outlook

## Impact Measurement

Automation should not be evaluated only by the amount of code produced.

Relevant indicators may include:
- Processing time before and after automation.
- Number of manual operations eliminated.
- Frequency of use.
- Errors detected or prevented.
- Time required for maintenance and exception handling.

Actual benefits should be measured in the operational environment rather than assumed.

## Data Protection

The examples in this repository should not contain confidential organizational information, personal data, patient data, or real business credentials.

Use synthetic or anonymized datasets when demonstrating the workflows.

## Development Status

These projects are practical automation solutions and experiments. Their suitability for use in other organizations depends on the underlying data, process requirements, business rules, and local validation.
