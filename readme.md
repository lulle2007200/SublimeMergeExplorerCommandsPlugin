# Sublime Merge Explorer Integration

**Sublime Merge Explorer Integration** provides context menu integration for the modern Windows Explorer.  

Currently, it adds the **"Open in Sublime Merge"** context menu entry for **directories tracked by git**.  
Future updates may expand its functionality.  

---

## Features

- Adds **"Open in Sublime Text"** to the Windows Explorer context menu for directories tracked by git.  
- By default, automatically checks for updates **every two weeks** (can be disabled in settings).  
- Simple installation and management of required components for Explorer integration.  

---

## Installation

After installing the package, you will be prompted to install the necessary components for Explorer integration.  

---

## Commands

This package provides the following commands:

### 1. `SublimeMergeExplorerIntegration: Check for Updates`
- Checks for available updates for the integration.  

### 2. `SublimeMergeExplorerIntegration: Install`
- Installs or reinstalls the Explorer integration components.  
- Useful if:
  - Your Sublime Merge installation directory has changed  
  - Installation was broken  

### 3. `SublimeMergeExplorerIntegration: Uninstall`
- Removes all components for Explorer integration.  

---

## Settings

- `auto_updates` If enabled, checks for updates on startup
- `update_check_interval` Interval for update checks in seconds. Two weeks by default.
- `installed_release_ts` Identifier of installed version. *DO NOT CHANGE MANUALLY*
- `last_update_check_ts` Timestamp of most recent update check. *DO NOT CHANGE MANUALLY*
- `last_update_check_ts` Path to sublime merge