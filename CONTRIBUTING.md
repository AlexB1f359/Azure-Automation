# Contributing to Azure Automation

Thank you for your interest in contributing to this project! We welcome improvements to the Azure Runbooks and supporting scripts.

## Getting Started

1.  **Fork the repository** on GitHub.
2.  **Clone your fork** locally.
3.  **Install Dependencies**: Ensure you have the required modules installed to test the scripts:
    ```powershell
    Install-Module -Name Microsoft.Graph -Scope CurrentUser
    Install-Module -Name ExchangeOnlineManagement -Scope CurrentUser
    ```

## Development Process

1.  Create a new branch for your feature or fix: `git checkout -b feature/amazing-feature`.
2.  Write your code. Please adhere to British English spelling in comments and documentation where possible.
3.  **Test your changes**:
    * Ensure scripts handle Azure Automation variables (like `Get-AutomationVariable`) gracefully if run locally (or use mock functions).
    * Verify that no hardcoded credentials or tenant IDs are included.
4.  Commit your changes with clear messages.
5.  Push to your fork and submit a Pull Request.

## PowerShell Style Guide

* Use `PascalCase` for variable names.
* Avoid using aliases in scripts (e.g., use `Where-Object` instead of `?`, and `ForEach-Object` instead of `%`).
* Ensure all new Runbooks include error handling (`Try`/`Catch` blocks) for external calls (Graph API, Exchange).
* If modifying `ABMMonitoring.ps1` or similar monitoring scripts, ensure the output logic (colours and HTML formatting) remains consistent.

## Reporting Bugs

Please use the Issue Templates provided to report bugs or request features.