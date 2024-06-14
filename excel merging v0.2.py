import pandas as pd
import msvcrt
from rich.progress import Progress, BarColumn, TimeRemainingColumn
from rich.console import Console
from colorama import Fore, Back, Style
from colorama import init
init(autoreset=True)
# import sys
import time

def getch():
    """Get a single character from the user without requiring Enter."""
    return msvcrt.getch().decode('utf-8')

def merge_worksheets(input_file, output_file, console):
    sheets_dict = pd.read_excel(input_file, sheet_name=None)
    merged_df = pd.DataFrame()
    
    # Ask user for each worksheet if they want to merge it or not
    with Progress("[progress.description]{task.description}", BarColumn(), "[progress.percentage]{task.percentage:>3.0f}%", TimeRemainingColumn()) as progress:
        task = progress.add_task("Worksheets covered percentage :", total=len(sheets_dict))
        for sheet_name, df in sheets_dict.items():
            print(f"Do you want to merge '{sheet_name}' worksheet? (y/n): ", end='', flush=True)
            merge_this_sheet = getch().lower()
            if merge_this_sheet == 'y':
                print(Fore.LIGHTGREEN_EX + f"Worksheet -> '{sheet_name}' selected for merging.")
                print("")
                time.sleep(0.5)
                merged_df = pd.concat([merged_df, df], ignore_index=True)
            elif merge_this_sheet == 'n':
                print(Fore.LIGHTYELLOW_EX + f"Skipping '{sheet_name}' worksheet.")
                print("")
            else:
                print(Fore.LIGHTRED_EX + "Invalid input. Please enter 'y' or 'n'.")
                print(Fore.LIGHTRED_EX + f"Worksheet -> '{sheet_name}' skipped due to wrong input.")
                print("")
            progress.update(task, advance=1)

    if merged_df.empty:
        print("No worksheets selected for merging. Exiting.")
        return
    with console.status("[bold magenta]Merging...") as status:
        with pd.ExcelWriter(output_file, engine='xlsxwriter') as writer:
            merged_df.to_excel(writer, index=False, sheet_name='Merged_Data')
            # Write original worksheets to the output file
            for sheet_name, df in sheets_dict.items():
                df.to_excel(writer, index=False, sheet_name=sheet_name)
        print(f"Merged worksheets saved to '{output_file}'")



if __name__ == "__main__":
    # input_file = r"C:\Users\91701\Desktop\Start-Excel-Test.xlsx"
    # output_file = r"C:\Users\91701\Desktop\Start-merged_output.xlsx"

    input_file = input("Enter file path :-  ")
    if '"' in input_file:
        inp = input_file.replace('"','')
    
    output_file = input_file.split(".")[0]+"_merged.xlsx"
    
    console = Console()
    merge_worksheets(input_file, output_file, console)