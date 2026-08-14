import pandas as pd
import os
import re
from datetime import datetime
import logging
from typing import Dict, List, Tuple
from openpyxl.styles import Font, PatternFill

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
#logger.addHandler(file_handler)
logger.addHandler(console_handler)

def clean_text_for_excel(text: str) -> str:
    """
    Clean text by removing illegal Excel characters and escape sequences.
    
    Args:
        text: Raw text to clean
        
    Returns:
        Cleaned text safe for Excel
    """
    if not text:
        return ""
    
    # Remove ANSI escape sequences (like \x1b[K)
    text = re.sub(r'\x1b\[[0-9;]*[a-zA-Z]', '', text)
    
    # Remove other control characters except newlines and tabs
    text = re.sub(r'[\x00-\x08\x0B\x0C\x0E-\x1F\x7F]', '', text)
    
    # Replace problematic characters with safe alternatives
    replacements = {
        '\x00': '',  # Null character
        '\x01': '',  # Start of heading
        '\x02': '',  # Start of text
        '\x03': '',  # End of text
        '\x04': '',  # End of transmission
        '\x05': '',  # Enquiry
        '\x06': '',  # Acknowledge
        '\x07': '',  # Bell
        '\x08': '',  # Backspace
        '\x0B': '',  # Vertical tab
        '\x0C': '',  # Form feed
        '\x0E': '',  # Shift out
        '\x0F': '',  # Shift in
        '\x10': '',  # Data link escape
        '\x11': '',  # Device control 1
        '\x12': '',  # Device control 2
        '\x13': '',  # Device control 3
        '\x14': '',  # Device control 4
        '\x15': '',  # Negative acknowledge
        '\x16': '',  # Synchronous idle
        '\x17': '',  # End of transmission block
        '\x18': '',  # Cancel
        '\x19': '',  # End of medium
        '\x1A': '',  # Substitute
        '\x1B': '',  # Escape
        '\x1C': '',  # File separator
        '\x1D': '',  # Group separator
        '\x1E': '',  # Record separator
        '\x1F': '',  # Unit separator
        '\x7F': '',  # Delete
    }
    
    for old_char, new_char in replacements.items():
        text = text.replace(old_char, new_char)
    
    # Remove any remaining non-printable characters
    text = ''.join(char for char in text if char.isprintable() or char in '\n\t\r')
    
    return text.strip()

class OutputProcessor:
    def __init__(self, output_directory: str = "./Output"):
        """
        Initialize the output processor.
        
        Args:
            output_directory: Directory containing the output files
        """
        self.output_directory = output_directory
        
    def find_output_files(self) -> List[str]:
        """Find all output files in the output directory that contain today's date."""
        output_files = []
        today_date = datetime.now().strftime('%Y-%m-%d')
        
        if not os.path.exists(self.output_directory):
            logger.error(f"Output directory not found: {self.output_directory}")
            return output_files
            
        for filename in os.listdir(self.output_directory):
            # Check if file meets all criteria:
            # 1. Ends with .txt
            # 2. Contains 'output' in filename
            # 3. Contains today's date
            if (filename.endswith('.txt') and 
                'output' in filename and 
                today_date in filename):
                output_files.append(os.path.join(self.output_directory, filename))
                
        logger.info(f"Found {len(output_files)} output files for today ({today_date})")
        return output_files
    
    def parse_output_file(self, file_path: str) -> Tuple[str, Dict[str, str]]:
        """
        Parse a single output file and extract host info and command outputs.
        
        Args:
            file_path: Path to the output file
            
        Returns:
            Tuple of (hostname, command_outputs_dict)
        """
        hostname = ""
        command_outputs = {}
        current_command = None
        current_output = []
        
        try:
            with open(file_path, 'r', encoding='utf-8') as f:
                lines = f.readlines()
                
            for line in lines:
                line = line.strip()
                
                # Extract hostname from header
                if line.startswith("SSH Command Output for Host:"):
                    hostname = line.split("SSH Command Output for Host:")[1].strip()
                    continue
                
                # Detect command sections
                if line.startswith("COMMAND") and ":" in line:
                    # Save previous command output if exists
                    if current_command and current_output:
                        command_outputs[current_command] = '\n'.join(current_output).strip()
                    
                    # Start new command
                    current_command = line.split(":", 1)[1].strip()
                    current_output = []
                    continue
                
                # Skip separator lines
                if line.startswith("=") or line.startswith("-") or not line:
                    continue
                
                # Skip summary lines
                if line.startswith("Total commands executed:") or line.startswith("Output file generated:"):
                    continue
                
                # Add line to current output
                if current_command:
                    current_output.append(line)
            
            # Save last command output
            if current_command and current_output:
                command_outputs[current_command] = '\n'.join(current_output).strip()
                
            logger.info(f"Parsed {len(command_outputs)} commands from {file_path}")
            return hostname, command_outputs
            
        except Exception as e:
            logger.error(f"Error parsing file {file_path}: {str(e)}")
            return "", {}
    
    def create_excel_summary(self, all_data: List[Tuple[str, Dict[str, str]]]) -> None:
        """
        Create Excel summary from parsed output data.
        
        Args:
            all_data: List of (hostname, command_outputs) tuples
        """
        if not all_data:
            logger.error("No data to process")
            return
            
        try:
            # Get all unique commands
            all_commands = set()
            for _, command_outputs in all_data:
                all_commands.update(command_outputs.keys())
            
            all_commands = sorted(list(all_commands))
            
            # Prepare Excel data
            excel_data = []
            
            # Create header
            header = ['Hostname'] + all_commands
            excel_data.append(header)
            
            # Add data rows
            for hostname, command_outputs in all_data:
                row = [hostname]
                
                for command in all_commands:
                    if command in command_outputs:
                        output = command_outputs[command]
                        # Clean the output for Excel compatibility
                        output = clean_text_for_excel(output)
                        # Truncate if too long for Excel
                        if len(output) > 32000:
                            output = output[:32000] + "\n[TRUNCATED - Output too long]"
                        row.append(output)
                    else:
                        row.append("Command not found")
                
                excel_data.append(row)
            
            # Create DataFrame
            df = pd.DataFrame(excel_data[1:], columns=excel_data[0])
            
            # Generate Excel filename
            timestamp = datetime.now().strftime('%Y-%m-%d_%H-%M-%S')
            excel_filename = f"./Output/ssh_summary_{timestamp}.xlsx"
            
            # Save to Excel with formatting
            with pd.ExcelWriter(excel_filename, engine='openpyxl') as writer:
                df.to_excel(writer, sheet_name='SSH_Outputs', index=False)
                
                # Get worksheet for formatting
                worksheet = writer.sheets['SSH_Outputs']
                
                # Auto-adjust column widths
                for column in worksheet.columns:
                    max_length = 0
                    column_letter = column[0].column_letter
                    
                    for cell in column:
                        try:
                            if len(str(cell.value)) > max_length:
                                max_length = len(str(cell.value))
                        except:
                            pass
                    
                    # Set column width (max 50 characters for readability)
                    adjusted_width = min(max_length + 2, 50)
                    worksheet.column_dimensions[column_letter].width = adjusted_width
                
                # Format header row
                header_font = Font(bold=True)
                header_fill = PatternFill(start_color="CCCCCC", end_color="CCCCCC", fill_type="solid")
                
                for cell in worksheet[1]:
                    cell.font = header_font
                    cell.fill = header_fill
            
            logger.info(f"Created Excel summary: {excel_filename}")
            
        except Exception as e:
            logger.error(f"Error creating Excel summary: {repr(e)}, line {e.__traceback__.tb_lineno}")
    
    def process_all_outputs(self) -> None:
        """Main method to process all output files and create Excel summary."""
        # Find output files
        output_files = self.find_output_files()
        
        if not output_files:
            logger.error("No output files found to process")
            return
        
        # Parse all output files
        all_data = []
        for file_path in output_files:
            hostname, command_outputs = self.parse_output_file(file_path)
            if hostname and command_outputs:
                all_data.append((hostname, command_outputs))
        
        if not all_data:
            logger.error("No valid data extracted from output files")
            return
        
        # Create Excel summary
        self.create_excel_summary(all_data)
        
        logger.info(f"Successfully processed {len(all_data)} hosts")

def main():
    """Main function to run the output processor."""
    processor = OutputProcessor()
    processor.process_all_outputs()

if __name__ == "__main__":
    main() 