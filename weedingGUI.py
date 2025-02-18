#!python
import tkinter as tk
from tkinter import ttk, scrolledtext
import sys
import os
from datetime import datetime
import threading
import queue
import re
import pandas as pd
import requests
import json
import html
from datetime import datetime, timedelta
from base64 import b64encode
from typing import Dict, List

# Constants
def load_credentials():
    """Load OCLC API credentials from credentials.dat"""
    try:
        with open('credentials.dat', 'r') as f:
            lines = f.readlines()
            if len(lines) >= 2:
                return lines[0].strip(), lines[1].strip()
            else:
                raise ValueError("credentials.dat must contain CLIENT_ID and CLIENT_SECRET on separate lines")
    except FileNotFoundError:
        raise FileNotFoundError("credentials.dat not found. Please create it with your CLIENT_ID and CLIENT_SECRET")
    except Exception as e:
        raise Exception(f"Error reading credentials: {str(e)}")

try:
    CLIENT_ID, CLIENT_SECRET = load_credentials()
except Exception as e:
    print(f"Error loading credentials: {str(e)}")
    CLIENT_ID = ""
    CLIENT_SECRET = ""


MAX_ENUMERATE = 3
MAX_DISTINCT = 5
MAX_COUNT = 250
MIN_COUNT = 50

class OCLCAuth:
    def __init__(self):
        self.client_id = CLIENT_ID
        self.secret = CLIENT_SECRET
        self.token = None
        self.token_expiry = None
        self.auth_endpoint = "https://oauth.oclc.org/token"

    def get_token(self):
        if self.token and datetime.now() < self.token_expiry:
            return self.token

        credentials = f"{self.client_id}:{self.secret}"
        encoded_credentials = b64encode(credentials.encode('utf-8')).decode('utf-8')

        headers = {
            'Authorization': f'Basic {encoded_credentials}',
            'Content-Type': 'application/x-www-form-urlencoded'
        }

        data = {
            'grant_type': 'client_credentials',
            'scope': 'wcapi:view_institution_holdings'
        }

        response = requests.post(self.auth_endpoint, headers=headers, data=data)
        if response.status_code == 200:
            token_data = response.json()
            self.token = token_data['access_token']
            self.token_expiry = datetime.now() + timedelta(seconds=token_data['expires_in'] - 60)
            return self.token
        else:
            raise Exception(f"Authentication failed: {response.text}")

class RedirectText:
    """Redirects stdout to GUI"""
    def __init__(self, text_widget: scrolledtext.ScrolledText, queue: queue.Queue):
        self.text_widget = text_widget
        self.queue = queue

    def write(self, string: str):
        self.queue.put(string)

    def flush(self):
        pass

class WeedingProcessor:
    """Handles the core weeding logic"""
    def __init__(self, cutoff_year: str):
        self.cutoff_year = cutoff_year
        
    def clean_hyperlink(self, value: str) -> str:
        """Clean Excel HYPERLINK formulas to extract just the text value."""
        if isinstance(value, str) and value.startswith('=HYPERLINK('):
            match = re.search(r'"([^"]+)"[^"]*$', value)
            return match.group(1) if match else value
        return value

    def process_initial_files(self, input_dir: str) -> pd.DataFrame:
        """Process initial Excel files and combine them."""
        xls_files = [f for f in os.listdir(input_dir) if f.endswith('.xls') and not f.endswith('.xls.zip')]
        
        if not xls_files:
            raise ValueError(f"No .xls files found in {input_dir}")
        
        dfs = []
        for filename in xls_files:
            file_path = os.path.join(input_dir, filename)
            try:
                df = pd.read_csv(file_path, sep='\t', skiprows=2, encoding='utf-8')
                if 'OCLC Number' in df.columns:
                    df['OCLC Number'] = df['OCLC Number'].apply(self.clean_hyperlink)
                dfs.append(df)
                print(f"Processed: {filename}")
            except Exception as e:
                print(f"Error processing {filename}: {str(e)}")
        
        if not dfs:
            raise ValueError("No files were successfully processed")
        
        return pd.concat(dfs, ignore_index=True)
    
    def sort_by_lcn(self, x: Dict) -> str:
        """Sort function for LC call numbers."""
        lcn = x["Local Call Number"]
        numstart = re.search("[A-Za-z]+", lcn).end()
        numlen = re.search("\D", lcn[numstart:]).start()
        return lcn[:numstart] + ("0" * (4-numlen)) + lcn[numstart:]

    def process_holdings(self, input_data: List[Dict], auth: OCLCAuth) -> List[List]:
        """Process holdings information using OCLC API."""
        api_base = "https://americas.discovery.api.oclc.org/worldcat/search/v2/bibs-holdings"
        processed_data = []

        for record in input_data:
            oclc_num = record["OCLC Number"]
            
            token = auth.get_token()
            headers = {
                'Authorization': f'Bearer {token}',
                'Accept': 'application/json'
            }
            
            params = {
                'oclcNumber': oclc_num,
                'holdingsAllEditions': False,
                'holdingsAllVariantRecords': False,
                'heldInState': 'US-FL',
                'limit': MAX_DISTINCT + 1
            }
            
            try:
                response = requests.get(api_base, headers=headers, params=params)
                response.raise_for_status()
                data_response = response.json()
                
                if 'briefRecords' in data_response:
                    holdings = data_response['briefRecords'][0]["institutionHolding"]["briefHoldings"]
                    inst_count = len(holdings)
                    
                    if inst_count == 1:
                        api_str = "Y"
                        institutions = ""
                    elif inst_count > 1 and inst_count <= MAX_ENUMERATE:
                        inst_names = [html.unescape(h["institutionName"]) for h in holdings[:MAX_ENUMERATE]]
                        if "Stetson University" in inst_names:
                            inst_names.remove("Stetson University")
                        api_str = f"1 of {inst_count}"
                        institutions = '; '.join(inst_names)
                    elif inst_count <= MAX_DISTINCT:
                        api_str = f"1 of {inst_count}"
                        institutions = ""
                    else:
                        api_str = "N"
                        institutions = ""
                else:
                    api_str = "Error"
                    institutions = ""
                    
            except requests.exceptions.RequestException as e:
                print(f"Error processing OCLC #{oclc_num}: {str(e)}")
                api_str = "Error"
                institutions = ""

            # Create output row
            row_data = [
                oclc_num, api_str, institutions, "",  # Empty string for In RCL?
                record["Title"], record["Author"], record["Publication Date"],
                record["Subject"], record["Format"], record["Edition"],
                record["Publisher"], record["Language"], record["LC Call Number"],
                record["Local Call Number"], record["Number of Circulations"],
                record["Last Circulated Date"]
            ]
            processed_data.append(row_data)
            
        return processed_data

    def run(self):
        """Main processing function"""
        try:
            # Create directories
            os.makedirs("input", exist_ok=True)
            os.makedirs("output", exist_ok=True)
            os.makedirs("output/xlsx files", exist_ok=True)

            print("Step 1: Processing initial files...")
            combined_df = self.process_initial_files("input")
            
            # Filter and process titles
            titles = []
            for _, row in combined_df.iterrows():
                data = row.to_dict()
                lastcirc = data["Last Circulated Date"]
                lastcircyear = 0 if "/" not in str(lastcirc) else str(lastcirc).split("/")[2]
                if (str(data["Publication Date"]) < self.cutoff_year or str(data["Publication Date"]) == "uuuu") and \
                   str(lastcircyear) < self.cutoff_year and data["Location"] == "FDSA Shelves":
                    titles.append(data)

            # Sort and group titles
            titles = sorted(titles, key=self.sort_by_lcn)
            sections = {}
            for title in titles:
                LCphrase = "".join([a if str.isalpha(a) else "" for a in str(title["LC Call Number"])[0:3]]).upper()
                title["Phrase"] = "UU" if str(title["Publication Date"]) == "uuuu" else LCphrase
                
                if title["Phrase"] not in sections:
                    sections[title["Phrase"]] = []
                sections[title["Phrase"]].append(title)

            # Separate into main and miscellaneous sections
            misc_titles = []
            main_sections = {}
            for phrase, phrase_titles in sections.items():
                if len(phrase_titles) < 20:
                    misc_titles.extend(phrase_titles)
                else:
                    main_sections[phrase] = phrase_titles

            print("\nStep 2: Creating intermediate files...")
            # Process main sections
            for section, titles in main_sections.items():
                numtitles = len(titles)
                numsections = round(numtitles/MAX_COUNT) if numtitles > MAX_COUNT else 1
                sectionlen = int(numtitles/numsections) + 1
                
                for i in range(numsections):
                    sectionnum = (" " + str(i+1)) if numsections > 1 else ""
                    filename = f"{section if section else 'blank'} candidates{sectionnum}.txt"
                    with open(os.path.join("output", filename), "w+", encoding="utf-8") as outfile:
                        headers = ["OCLC Number", "Only Lib?", "Others Holding", "In RCL?", "Title", "Author", 
                                  "Publication Date", "Subject", "Format", "Edition", "Publisher", "Language", 
                                  "LC Call Number", "Local Call Number", "Number of Circulations", "Last Circulated Date"]
                        outfile.write("\t".join(headers) + '\n')
                        for title in titles[i*sectionlen:(i+1)*sectionlen]:
                            out = "\t".join([str(title.get(field, "")) for field in headers])
                            outfile.write(out + '\n')

            # Process miscellaneous section
            if misc_titles:
                misc_titles = sorted(misc_titles, key=self.sort_by_lcn)
                with open(os.path.join("output", "MISC candidates.txt"), "w+", encoding="utf-8") as outfile:
                    headers = ["OCLC Number", "Only Lib?", "Others Holding", "In RCL?", "Title", "Author", 
                              "Publication Date", "Subject", "Format", "Edition", "Publisher", "Language", 
                              "LC Call Number", "Local Call Number", "Number of Circulations", "Last Circulated Date"]
                    outfile.write("\t".join(headers) + '\n')
                    for title in misc_titles:
                        out = "\t".join([str(title.get(field, "")) for field in headers])
                        outfile.write(out + '\n')

            print("\nStep 3: Processing OCLC holdings...")
            auth = OCLCAuth()
            # Process each intermediate file
            for filename in os.listdir("output"):
                if filename.endswith("candidates.txt"):
                    file_prefix = filename.replace(" candidates.txt", "")
                    input_file = os.path.join("output", filename)
                    output_file = os.path.join("output/xlsx files", f"{file_prefix} candidates.xlsx")
                    
                    try:
                        print(f"Processing {filename}...")
                        # Read the intermediate file
                        df = pd.read_csv(input_file, sep='\t', encoding='utf-8')
                        # Process holdings
                        processed_data = self.process_holdings(df.to_dict('records'), auth)
                        # Create output DataFrame
                        columns = ['OCLCNum', 'Only Lib?', 'Others Holding', 'In RCL?', 'Title', 'Author', 
                                  'Publication Date', 'Subject', 'Format', 'Edition', 'Publisher', 'Language', 
                                  'LC Call Number', 'Local Call Number', 'Number of Circulations', 'Last Circulated Date']
                        result_df = pd.DataFrame(processed_data, columns=columns)
                        # Write to Excel
                        result_df.to_excel(output_file, index=False)
                        print(f"Results written to {output_file}")
                    except Exception as e:
                        print(f"An error occurred processing {filename}: {str(e)}")
                        continue

            print("\nProcess complete! Check the 'output/xlsx files' directory for results.")
            return True

        except Exception as e:
            print(f"\nError occurred: {str(e)}")
            return False

class WeedingGUI:
    def __init__(self, root: tk.Tk):
        self.root = root
        self.root.title("Library Weeding Process")
        self.root.geometry("800x600")
        
        # Create message queue
        self.msg_queue = queue.Queue()
        
        # Create main frame
        main_frame = ttk.Frame(root, padding="10")
        main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        # Configure grid
        root.columnconfigure(0, weight=1)
        root.rowconfigure(0, weight=1)
        main_frame.columnconfigure(1, weight=1)
        main_frame.rowconfigure(2, weight=1)

        # Cutoff year selection
        year_frame = ttk.Frame(main_frame)
        year_frame.grid(row=0, column=0, columnspan=2, pady=(0, 10), sticky=tk.W)
        
        ttk.Label(year_frame, text="Cutoff Year:").grid(row=0, column=0, padx=(0, 10))
        
        current_year = datetime.now().year
        years = list(range(current_year, 1900, -1))
        self.year_var = tk.StringVar(value="2004")
        year_combo = ttk.Combobox(year_frame, textvariable=self.year_var, values=years, width=10)
        year_combo.grid(row=0, column=1)

        # Buttons frame
        button_frame = ttk.Frame(main_frame)
        button_frame.grid(row=1, column=0, columnspan=2, pady=(0, 10))
        
        # Process button
        self.process_button = ttk.Button(button_frame, text="Start Processing", command=self.start_processing)
        self.process_button.grid(row=0, column=0, padx=5)
        
        # Help button
        self.help_button = ttk.Button(button_frame, text="Help", command=self.show_help)
        self.help_button.grid(row=0, column=1, padx=5)

        # Progress indicator
        self.progress_var = tk.StringVar(value="Ready to process")
        self.progress_label = ttk.Label(main_frame, textvariable=self.progress_var)
        self.progress_label.grid(row=3, column=0, columnspan=2, pady=(5, 0))

        # Output text area
        self.output_text = scrolledtext.ScrolledText(main_frame, height=20, width=80)
        self.output_text.grid(row=2, column=0, columnspan=2, sticky=(tk.W, tk.E, tk.N, tk.S))

        # Redirect stdout
        sys.stdout = RedirectText(self.output_text, self.msg_queue)

        # Start message checking
        self.check_msg_queue()

    def check_msg_queue(self):
        """Check for new messages and update text widget"""
        while True:
            try:
                msg = self.msg_queue.get_nowait()
                self.output_text.insert(tk.END, msg)
                self.output_text.see(tk.END)
                self.output_text.update_idletasks()
            except queue.Empty:
                break
        self.root.after(100, self.check_msg_queue)

    def start_processing(self):
        """Start processing in separate thread"""
        self.process_button.configure(state='disabled')
        self.progress_var.set("Processing... Please wait")
        self.output_text.delete(1.0, tk.END)
        
        process_thread = threading.Thread(target=self.run_process)
        process_thread.daemon = True
        process_thread.start()

    def run_process(self):
        """Run the main processing logic"""
        processor = WeedingProcessor(self.year_var.get())
        success = processor.run()
        
        if success:
            self.root.after(0, self.processing_complete)
        else:
            self.root.after(0, self.processing_error)

    def processing_complete(self):
        """Update GUI after successful processing"""
        self.process_button.configure(state='normal')
        self.progress_var.set("Processing complete!")

    def processing_error(self):
        """Update GUI after processing error"""
        self.process_button.configure(state='normal')
        self.progress_var.set("Error occurred during processing")

    def show_help(self):
        """Show help dialog with README information"""
        help_window = tk.Toplevel(self.root)
        help_window.title("Library Weeding Process - Help")
        help_window.geometry("600x400")
        
        # Make the window modal
        help_window.transient(self.root)
        help_window.grab_set()
        
        # Add scrolled text widget for help content
        help_text = scrolledtext.ScrolledText(help_window, wrap=tk.WORD, padx=10, pady=10)
        help_text.pack(expand=True, fill='both')
        
        readme_content = """Library Weeding Process - README

Prerequisites:
1. Directory Structure
   - An 'input' directory must exist in the same location as the program
   - Place all .xls files to be processed in the 'input' directory
   - The program will create 'output' and 'output/xlsx files' directories automatically

2. Credentials
   - A 'credentials.dat' file must exist in the same location as the program
   - The file should contain two lines:
     Line 1: Your OCLC API Client ID (80 characters long)
     Line 2: Your OCLC API Client Secret (32 characters long)
   - These values should be accessible at https://platform.worldcat.org/wskey/
   - Contact your system administrator if you need these credentials

3. Input Files
   - This program is designed to work with unzipped output from WMS.
   - Only .xls files are processed (not .xlsx or .xls.zip)
   - These files are actually tab-delimited text files that WMS suffixes with .xls.
   - Files should contain the following columns:
     * OCLC Number
     * Title
     * Author
     * Publication Date
     * Subject
     * Format
     * Edition
     * Publisher
     * Language
     * LC Call Number
     * Local Call Number
     * Number of Circulations
     * Last Circulated Date
     * Location

Usage:
1. Select the desired cutoff year from the dropdown
2. Click 'Start Processing' to begin
3. Watch the progress in the main window
4. When complete, check the 'output/xlsx files' directory for results

If you encounter any errors:
1. Verify all prerequisites are met
2. Check that your credentials are valid
3. Ensure your input files are in the correct format
4. Contact technical support if problems persist"""
        
        help_text.insert('1.0', readme_content)
        help_text.configure(state='disabled')  # Make text read-only
        
        # Add close button
        close_button = ttk.Button(help_window, text="Close", command=help_window.destroy)
        close_button.pack(pady=10)

def main():
    root = tk.Tk()
    app = WeedingGUI(root)
    root.mainloop()

if __name__ == "__main__":
    main()