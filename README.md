# Fix-WindowPositions

A PowerShell script to restore misplaced or off-screen windows back into the visible desktop area.  
This script is helpful when windows are stuck outside of the visible screen (common with multi-monitor setups).

## Features
- Detects windows that are off-screen or positioned incorrectly.
- Moves them back to the primary screen for visibility.
- Provides a dry-run option to preview changes without applying them.

## Usage
1. Open PowerShell with administrator privileges.
2. Run the script with:
   ```powershell
   .\Fix-WindowPositions.ps1
   ```

   To test without making changes:
   ```powershell
   .\Fix-WindowPositions.ps1 -DryRun
   ```

## Notes
- Useful for multi-monitor setups where windows can get lost outside the visible area.
- Can be customized further depending on your workflow.

## License
This project is open-source under the MIT License.
