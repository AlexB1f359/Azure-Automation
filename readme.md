
# Azure Automation

This repository contains PowerShell scripts and runbooks for automating Azure resource monitoring and management tasks.

## Contents

### Azure Runbooks
- **ABMMonitoring.ps1** - Monitors Apple certificates ands tokens used within Intune
- **CertificateAndSecretMonitoring.ps1** - Monitors certificates and secrets for home tenant Enterprise apps and service principals.

### Supporting Scripts
- **ApplyManagedIdentityPermissions.ps1** - Configures and applies role-based access control (RBAC) permissions for managed identities
- **ExchangeApplicationAccessPolicy.ps1** - Manages Exchange application access policies

## Usage

These scripts are designed to run in Azure Automation accounts. Deploy the runbooks to your Azure Automation account and configure them with the appropriate managed identities and permissions using the supporting scripts.

## Requirements

- Azure PowerShell modules
- Appropriate Azure permissions and role assignments
- Graph PowerShell modules
