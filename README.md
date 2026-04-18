# Microsoft Excel (microsoft-excel)
APIs for automating, integrating, and extending Microsoft Excel functionality including workbook management, data manipulation, charting, and formula execution through Microsoft Graph REST APIs.

**URL:** [Visit APIs.json URL](https://raw.githubusercontent.com/api-evangelist/microsoft-excel/refs/heads/main/apis.yml)

**Run:** [Capabilities Using Naftiko](https://github.com/naftiko/fleet?utm_source=api-evangelist&utm_medium=readme&utm_campaign=company-api-evangelist&utm_content=repo)

## Tags:

 - Automation, Data Analysis, Microsoft, Microsoft 365, Office, Spreadsheets

## Timestamps

- **Created:** 2024
- **Modified:** 2026-04-18

## APIs

### Microsoft Graph Excel API
REST API for accessing and manipulating Excel workbooks stored in OneDrive for Business, SharePoint sites, or Group drives through Microsoft Graph. Supports worksheet, table, chart, range, named item, and function operations.

**Human URL:** [https://learn.microsoft.com/en-us/graph/api/resources/excel](https://learn.microsoft.com/en-us/graph/api/resources/excel)

#### Tags:

 - Cloud, Microsoft Graph, OneDrive, REST API, SharePoint

#### Properties

- [Documentation](https://learn.microsoft.com/en-us/graph/api/resources/excel)
- [OpenAPI](openapi/microsoft-excel-graph-api.yaml)
- [Authentication](https://learn.microsoft.com/en-us/graph/auth/)
- [APIReference](https://learn.microsoft.com/en-us/graph/api/resources/excel)
- [SDK](https://learn.microsoft.com/en-us/graph/sdks/sdks-overview)

### Excel JavaScript API
Office Add-ins API for building custom functionality within Excel using JavaScript and TypeScript.

**Human URL:** [https://learn.microsoft.com/en-us/office/dev/add-ins/reference/overview/excel-add-ins-reference-overview](https://learn.microsoft.com/en-us/office/dev/add-ins/reference/overview/excel-add-ins-reference-overview)

#### Tags:

 - Client-Side, JavaScript, Office Add-Ins

#### Properties

- [Documentation](https://learn.microsoft.com/en-us/javascript/api/excel)
- [GettingStarted](https://learn.microsoft.com/en-us/office/dev/add-ins/quickstarts/excel-quickstart-jquery)

### Office Scripts API
TypeScript-based automation for Excel on the web, enabling Power Automate integration for business process automation.

**Human URL:** [https://learn.microsoft.com/en-us/office/dev/scripts/overview/excel](https://learn.microsoft.com/en-us/office/dev/scripts/overview/excel)

#### Tags:

 - Automation, Excel Online, Power Automate, TypeScript

#### Properties

- [Documentation](https://learn.microsoft.com/en-us/office/dev/scripts/develop/scripting-fundamentals)
- [CodeExamples](https://learn.microsoft.com/en-us/office/dev/scripts/resources/samples/samples-overview)
- [Tutorials](https://learn.microsoft.com/en-us/office/dev/scripts/tutorials/excel-tutorial)

## Common Properties

- [DeveloperPortal](https://developer.microsoft.com/en-us/microsoft-365)
- [Support](https://support.microsoft.com/excel)
- [TermsOfService](https://www.microsoft.com/en-us/legal/terms-of-use)
- [PrivacyPolicy](https://privacy.microsoft.com/)
- [Pricing](https://www.microsoft.com/en-us/microsoft-365/excel)
- [Blog](https://techcommunity.microsoft.com/t5/excel-blog/bg-p/ExcelBlog)
- [GitHubOrganization](https://github.com/OfficeDev)

## Features

| Name | Description |
|------|-------------|
| Workbook Sessions | Persistent and non-persistent sessions for efficient API operations. |
| Cell Range Operations | Read, write, and format individual cells or ranges of cells. |
| Table Management | Create, update, sort, and filter structured tables. |
| Chart Generation | Create and manage charts with customizable source data. |
| Workbook Functions | Invoke Excel functions programmatically via the API. |
| Named Ranges | Define and manage named ranges for reusable cell references. |

## Use Cases

| Name | Description |
|------|-------------|
| Automated Reporting | Generate and update Excel reports from external data sources. |
| Data Entry Automation | Programmatically insert and update data in spreadsheets. |
| Financial Modeling | Use Excel functions API for financial calculations. |
| Dashboard Generation | Create charts and visualizations from data. |

## Integrations

| Name | Description |
|------|-------------|
| Microsoft Power Automate | Automate Excel workflows using Power Automate connectors. |
| SharePoint | Access and modify Excel files stored in SharePoint sites. |
| OneDrive | Work with Excel workbooks stored in OneDrive for Business. |
| Microsoft Teams | Share and collaborate on Excel workbooks within Teams. |

## Artifacts

Machine-readable API specifications organized by format.

### OpenAPI

- [Microsoft Graph Excel API](openapi/microsoft-excel-graph-api.yaml)

### JSON Schema

- [Worksheet](json-schema/excel-graph-api-worksheet-schema.json)
- [Table](json-schema/excel-graph-api-table-schema.json)
- [TableRow](json-schema/excel-graph-api-table-row-schema.json)
- [TableColumn](json-schema/excel-graph-api-table-column-schema.json)
- [Chart](json-schema/excel-graph-api-chart-schema.json)
- [Range](json-schema/excel-graph-api-range-schema.json)
- [NamedItem](json-schema/excel-graph-api-named-item-schema.json)

### JSON Structure

- [Worksheet](json-structure/excel-graph-api-worksheet-structure.json)
- [Table](json-structure/excel-graph-api-table-structure.json)
- [TableRow](json-structure/excel-graph-api-table-row-structure.json)
- [TableColumn](json-structure/excel-graph-api-table-column-structure.json)
- [Chart](json-structure/excel-graph-api-chart-structure.json)
- [Range](json-structure/excel-graph-api-range-structure.json)
- [NamedItem](json-structure/excel-graph-api-named-item-structure.json)

### JSON-LD

- [Microsoft Excel Graph API Context](json-ld/microsoft-excel-graph-api-context.jsonld)

### Examples

- [Worksheet](examples/excel-graph-api-worksheet-example.json)
- [Table](examples/excel-graph-api-table-example.json)
- [TableRow](examples/excel-graph-api-table-row-example.json)
- [TableColumn](examples/excel-graph-api-table-column-example.json)
- [Chart](examples/excel-graph-api-chart-example.json)
- [Range](examples/excel-graph-api-range-example.json)
- [NamedItem](examples/excel-graph-api-named-item-example.json)

## Capabilities

Naftiko capabilities organized as shared per-API definitions composed into customer-facing workflows.

### Shared Per-API Definitions

- [Microsoft Graph Excel API](capabilities/shared/excel-graph-api.yaml) -- 7 operations for workbook management

### Workflow Capabilities

| Workflow | APIs Combined | Tools | Persona |
|----------|--------------|-------|---------|
| [Spreadsheet Automation](capabilities/spreadsheet-automation.yaml) | Excel Graph API | 7 | Data Analyst, Business Analyst |

## Vocabulary

- [Microsoft Excel Vocabulary](vocabulary/microsoft-excel-vocabulary.yaml) -- Unified taxonomy mapping 9 resources, 9 actions, 1 workflow, and 2 personas across operational (OpenAPI) and capability (Naftiko) dimensions

## Rules

- [Microsoft Excel Spectral Rules](rules/microsoft-excel-spectral-rules.yml) -- 22 rules across 9 categories enforcing Microsoft Excel API conventions

## Maintainers

**FN:** Kin Lane

**Email:** kin@apievangelist.com
