# SSH Network Device Automation Tool with Excel config

A comprehensive Python tool for automating SSH connections to network devices, executing commands from excel, and generating organized excel output reports.

## Overview

This tool consists of two main components:
1. **SSH_Auto_Extract.py** - Connects to network devices via SSH and executes commands
2. **ssh_output_to_excel.py** - Processes output files and creates Excel summaries

## Features

- **Multi-device SSH automation** with Excel-based configuration
- **Command execution from Excel sheet** 
- **Structured output** with detailed logging
- **Excel summary generation** for easy comparison
- **Error handling** with line number tracking
- **Flexible configuration** via Excel files

## Requirements

### Python Dependencies
```bash
pip install -r requirements.txt
```

### Required Packages
- `paramiko>=2.8.0` - SSH connections
- `pandas>=1.3.0` - Excel file handling
- `openpyxl>=3.0.0` - Modern Excel format support
- `xlrd>=2.0.0` - Legacy Excel format support
- `odfpy>=1.4.0` - OpenDocument format support

## Configuration

### Excel File Structure (`hosts.xlsx`)

**Sheet1 (Hosts Configuration):**
| host_name | username | password |
|-----------|----------|----------|
| 172.30.10.1 | root | rootpw01 |
| 172.30.10.2 | admin | admin123 |

**Sheet2 (Commands):**
| Commands |
|----------|
| ifconfig |
| cat /etc/os-release |


## Usage

### Step 1: SSH Automation
```bash
python SSH_Auto_Extract.py
```

This script will:
- Read host configurations from Sheet1 of `hosts.xlsx`
- Read commands from Sheet2 of `hosts.xlsx`
- Connect to each host via SSH
- Execute all commands on each device
- Generate individual output files for each host in `./Output` directory
- **NOTE: the script reads the '#' symbol as the indicator to start the next command, please adjust accordto your connnected device**
- add `--no-exec` argument if command execution and text file output is not needed (e.g. for testing connection)

### Step 2: Excel Summary Generation
```bash
python ssh_output_to_excel.py
```

This script will:
- Read all output files from the `./Output` directory
- Parse command outputs and host information
- Create a comprehensive Excel summary
- Generate formatted Excel file with hosts as rows and commands as columns

## Output Files

### Individual Host Files
- Location: `./Output/`
- Format: `host1_172.30.10.1_output_YYYY-MM-DD.txt`
- Content: All command outputs for each host, IP with timestamps

### Excel Summary File
- Location: `./Output/`
- Format: `ssh_summary_YYYY-MM-DD_HH-MM-SS.xlsx`
- Content: Matrix view with hosts as rows and command outputs as columns

## File Structure

```
project/
├── SSH_Auto_Extract.py      # Main SSH automation script
├── ssh_output_to_excel.py       # Excel summary generator
├── hosts.xlsx              # Configuration file (hosts + commands)
├── requirements.txt        # Python dependencies
├── README.md              # This file
├── Output/                # Generated output files
│   ├── host1_*.txt       # Individual host outputs
│   ├── host2_*.txt       # Individual host outputs
│   └── ssh_summary_*.xlsx # Excel summary
└── ssh_automation.log     # Execution log file
```



## Security Notes

- **Passwords are stored in plain text** in the Excel file
- **Consider using SSH keys** for production environments
- **Restrict file permissions** on configuration files
- **Use secure networks** for SSH connections
- **The script reads the '#' symbol as the indicator to start the next command**

## Troubleshooting

### Common Issues

1. **"File is not a zip file"**
   - Ensure Excel file is properly formatted
   - Try recreating the Excel file

2. **"Connection failed"**
   - Verify host IP addresses
   - Check username/password
   - Ensure network connectivity

3. **"Missing required columns"**
   - Verify Excel sheet structure
   - Check column names match exactly

4. **"IllegalCharacterError"**
   - Script automatically handles this
   - If persistent, check command outputs
   - 
### Current BUG

1. The Excel generated might be wrong if there are repeated command

### Debug Mode

Add debug prints to see loaded configurations:
```python
print(COMMANDS)  # Shows loaded commands
print(hosts_config)  # Shows loaded hosts
```
## Version History

- **v1.0.0** - Initial release with SSH automation and Excel output
- **v1.1.0** - Added Excel-based configuration
- **v1.2.0** - Enhanced error handling and logging
- **v1.3.0** - Added character encoding fixes for Excel compatibility 
