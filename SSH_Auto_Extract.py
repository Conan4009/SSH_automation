import paramiko
from typing import Dict, List, Union
import json
import os
from datetime import datetime
import logging
import time
import pandas as pd

# Configure logging
logger = logging.getLogger()
logger.setLevel(logging.INFO)

# Create a file handler
file_handler = logging.FileHandler('./ssh_automation.log')
file_handler.setLevel(logging.INFO)

# Create a console handler
console_handler = logging.StreamHandler()
console_handler.setLevel(logging.INFO)

# Create a formatter and set it for both handlers
formatter = logging.Formatter('%(asctime)s - %(levelname)s - %(message)s')
file_handler.setFormatter(formatter)
console_handler.setFormatter(formatter)

# Add the handlers to the logger
logger.addHandler(file_handler)
logger.addHandler(console_handler)

import argparse

# Create a parser object
parser = argparse.ArgumentParser(description="exec flag to enable/disable command execution")

# Add the -exec argument, expecting a string value (like "true" or "false")
parser.add_argument('--no-exec',dest='no_exec', action='store_true', help='Disable command execution')

# Parse the arguments
args = parser.parse_args()



def load_commands_from_excel(file_path: str = 'hosts.xlsx', sheet_name: str = 'Sheet2') -> List[str]:
    """
    Load commands from sheet 2 of the Excel file.
    
    Args:
        file_path: Path to the Excel file
        sheet_name: Name of the sheet containing commands (default: Sheet2)
        
    Returns:
        List of commands
    """
    commands = []
    try:
        # Read commands from sheet 2 of the Excel file
        try:
            df_commands = pd.read_excel(file_path, sheet_name=sheet_name, engine='openpyxl')
        except Exception as e1:
            logger.warning(f"Failed to read commands with openpyxl: {str(e1)}")
            try:
                df_commands = pd.read_excel(file_path, sheet_name=sheet_name, engine='xlrd')
            except Exception as e2:
                logger.warning(f"Failed to read commands with xlrd: {str(e2)}")
                try:
                    df_commands = pd.read_excel(file_path, sheet_name=sheet_name, engine='odf')
                except Exception as e3:
                    logger.error(f"Failed to read commands from Excel sheet. Errors: openpyxl={str(e1)}, xlrd={str(e2)}, odf={str(e3)}, line {e3.__traceback__.tb_lineno}")
                    raise Exception(f"Cannot read commands from Excel sheet {sheet_name}")
        
        # Extract commands from the first column
        if not df_commands.empty:
            # Get the first column (assuming commands are in the first column)
            first_column = df_commands.iloc[:, 0]
            for command in first_column:
                if pd.notna(command):  # Check if not NaN
                    command_str = str(command).strip()
                    if command_str and not command_str.startswith('#'):
                        commands.append(command_str)
        
        logger.info(f"Loaded {len(commands)} commands from {file_path} sheet '{sheet_name}'")
        return commands
        
    except FileNotFoundError as e:
        logger.error(f"Excel file not found: {file_path}, line {e.__traceback__.tb_lineno}")
        # Return default commands if file not found
        return [
            'show module',
        ]
    except Exception as e:
        logger.error(f"Error reading commands from Excel: {str(e)}, line {e.__traceback__.tb_lineno}")
        # Return default commands if error occurs
        return [
            'show module',
        ]

def load_hosts_from_excel(commands,file_path: str = 'hosts.xlsx') -> List[Dict]:
    """
    Load host configurations from an Excel file.
    
    Args:
        file_path: Path to the Excel file
        
    Returns:
        List of host configuration dictionaries
    """
    try:
        # Read Excel file with explicit engine specification
        try:
            df = pd.read_excel(file_path, engine='openpyxl')
        except Exception as e1:
            logger.warning(f"Failed to read with openpyxl: {str(e1)}")
            try:
                # Try xlrd for older .xls files
                df = pd.read_excel(file_path, engine='xlrd')
            except Exception as e2:
                logger.warning(f"Failed to read with xlrd: {str(e2)}")
                try:
                    # Try odf for .ods files
                    df = pd.read_excel(file_path, engine='odf')
                except Exception as e3:
                    logger.error(f"Failed to read Excel file with all engines. Errors: openpyxl={str(e1)}, xlrd={str(e2)}, odf={str(e3)}, line {e3.__traceback__.tb_lineno}")
                    raise Exception(f"Cannot read Excel file {file_path} with any supported engine")
        
        # Validate required columns
        required_columns = ['host_name', 'username', 'password']
        missing_columns = [col for col in required_columns if col not in df.columns]
        
        if missing_columns:
            logger.error(f"Missing required columns in Excel file: {missing_columns}, line {e.__traceback__.tb_lineno}")
            logger.info(f"Available columns: {list(df.columns)}")
            return []
        
        hosts = []
        for index, row in df.iterrows():
            host_config = {
                'hostname': str(row['host_name']),
                'username': str(row['username']),
                'password': str(row['password']),
                'commands': commands,
                'output_file': f'./Output/host{index+1}_{str(row['host_name'])}_output_{datetime.now().strftime('%Y-%m-%d')}.txt'
            }
            hosts.append(host_config)
            
        logger.info(f"Loaded {len(hosts)} host configurations from {file_path}")
        return hosts
        
    except FileNotFoundError as e:
        logger.error(f"Excel file not found: {file_path}, line {e.__traceback__.tb_lineno}")
        # Return default configuration if file not found
        return []
    except Exception as e:
        logger.error(f"Error reading Excel file: {str(e)}, line {e.__traceback__.tb_lineno}")
        return []

class SSHAutomation:
    def __init__(self, config_path: str = None, prompt: str = '#'):
        """
        Initialize SSHAutomation with optional configuration path.
        
        Args:
            config_path: Path to JSON configuration file
        """
        self.config_path = config_path
        self.ssh_client = None
        self.shell = None
        self.prompt = prompt
        
        # Default configuration structure
        self.default_config = {
            "hosts": []
        }
        
        # Load configuration from file if provided
        if config_path:
            self.load_config()

    def load_config(self) -> None:
        """Load configuration from JSON file. To be updated"""
        try:
            with open(self.config_path, 'r') as f:
                self.config = json.load(f)
        except FileNotFoundError as e:
            logger.error(f"Configuration file not found: {self.config_path}, line {e.__traceback__.tb_lineno}")
            raise
        except json.JSONDecodeError as e:
            logger.error(f"Invalid JSON format in configuration file, line {e.__traceback__.tb_lineno}")
            raise

    def connect(self, host: Dict[str, str]) -> bool:
        """Establish SSH connection to host."""
        try:
            self.ssh_client = paramiko.SSHClient()
            self.ssh_client.set_missing_host_key_policy(paramiko.AutoAddPolicy())
            
            logger.info(f"Connecting to {host['hostname']}")
            self.ssh_client.connect(
                hostname=host['hostname'],
                username=host['username'],
                password=host['password']
            )
            logger.info(f"Successfully connected to {host['hostname']}")
            
            # Invoke shell with larger buffer size for better output handling
            self.shell = self.ssh_client.invoke_shell(width=1000, height=1000)
            time.sleep(3)
            if self.shell.recv_ready():
                self.shell.recv(9999)
            time.sleep(0.3)
            return True
        except Exception as e:
            logger.error(f"Connection failed to {host['hostname']}: {str(e)}, line {e.__traceback__.tb_lineno}")
            return False

    def execute_commands(self, host: Dict[str, Union[str, List[str]]]) -> Dict[str, str]:
        """Execute multiple commands on remote host using shell."""
        results = {}
        
        if not isinstance(host.get('commands', []), list):
            host['commands'] = [host.get('command', '')]
        
        for command in host['commands']:
            try:
                time.sleep(0.5)
                if self.shell.recv_ready():
                    self.shell.recv(1024)
                    time.sleep(0.3)
                # Send command and wait for prompt
                self.shell.send(command + '\n')
                logger.info(f"Sent command: {command}")
                time.sleep(0.3)
                # Wait for command to complete and collect output
                output = self._receive_output(self.prompt)
                
                # Store results
                results[command] = output
                
                logger.info(f"Executed command '{command}' on {host['hostname']}")
                
            except Exception as e:
                logger.error(f"Failed to execute command '{command}': {str(e)}, line {e.__traceback__.tb_lineno}")
                results[command] = f"Error: {str(e)}"
        
        return results

    def _receive_output(self,prompt: str='#') -> str:
        """Receive output from shell until prompt is detected."""
        output = ''
        
        # Wait for output
        start_time = time.time()
        while True:
            if self.shell.recv_ready():
                chunk = self.shell.recv(1024).decode()
                output += chunk
                print(chunk)
                # Check if we've received the prompt
                if chunk.strip().endswith(prompt):
                    break
                    
            # Check for timeout
            if (time.time() - start_time > 3) and (time.time() - start_time < 60):  # 30 second timeout
                self.shell.send(' ')
                logger.info('sending space')
                time.sleep(5)
            elif time.time() - start_time > 60:  # 60 second timeout
                logger.warning("Command execution timed out")
                break
                
            # Small delay to avoid busy-waiting
            time.sleep(0.1)
            
        return output.strip()

    def save_output(self, host: Dict[str, Union[str, List[str]]], results: Dict[str, str]) -> None:
        """Save all command outputs to a single file per host."""
        filename = host.get('output_file', 
                           f"./Output/{host['hostname']}_output_{datetime.now().strftime('%Y-%m-%d')}.txt")
        
        try:
            with open(filename, 'w', encoding='utf-8') as f:
                # Write header with host information
                f.write(f"SSH Command Output for Host: {host['hostname']}\n")
                f.write(f"Date: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n")
                f.write("=" * 80 + "\n\n")
                
                # Write each command and its output
                for i, (command, result) in enumerate(results.items(), 1):
                    f.write(f"COMMAND {i}: {command}\n")
                    f.write("-" * 60 + "\n")
                    f.write(result)
                    f.write("\n\n")
                    f.write("=" * 80 + "\n\n")
                
                # Write summary
                f.write(f"Total commands executed: {len(results)}\n")
                f.write(f"Output file generated: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n")
            
            logger.info(f"Saved all outputs to {filename}")
        except IOError as e:
            logger.error(f"Failed to save output to {filename}: {str(e)}, line {e.__traceback__.tb_lineno}")

    def process_host(self, host: Dict[str, Union[str, List[str]]],no_exec=False) -> None:
        """Process a single host configuration."""
        if not self.connect(host):
            return
        
        try:
            if not no_exec:
                results = self.execute_commands(host)
                self.save_output(host, results)
        finally:
            if self.shell:
                self.shell.close()
            if self.ssh_client:
                self.ssh_client.close()
                logger.info(f"Connection closed for {host['hostname']}")

def main():
    """Main function to demonstrate usage."""
    # Load commands from file
    COMMANDS = load_commands_from_excel()
    # Load hosts configuration from Excel
    hosts_config = load_hosts_from_excel(COMMANDS)
    
    # Create SSH automation instance
    ssh_automation = SSHAutomation()
    #ssh_automation = SSHAutomation(config_path)
    # Process each host
    for host in hosts_config:
        ssh_automation.process_host(host,args.no_exec)
        print('Host completed, Stopping 5 sec')
        time.sleep(5)

if __name__ == "__main__":
    main()
