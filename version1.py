
import mouse
import pyautogui
import time
import pyperclip
import subprocess
import psutil
import os
from module1.textclip import clipcopy as clipcopy
from pynput import keyboard

## Enable PyAutoGUI's built-in failsafe
pyautogui.FAILSAFE = True  # Move the mouse to the top-left corner to stop the script
# Screen dimensions
screen_width, screen_height = pyautogui.size()
taskbar_height = 40  # Adjust based on your taskbar size
taskbar_region = (0, screen_height - taskbar_height, screen_width, taskbar_height)
excel_img = "excel.png"
# Path to the chrome icon image
chrome_img = "chrome.png"

lockpage = "lockpage.png"
def duplicate_and_rename_file_terminal(template_path, new_file_name, output_directory):
    """
    Duplicate a file and rename it using terminal commands.
    Args:
        template_path (str): The path to the template file.
        new_file_name (str): The new name for the duplicated file.
        output_directory (str): The directory where the new file will be saved.
    """
    try:
        # Ensure the output directory exists
        os.makedirs(output_directory, exist_ok=True)

        # Construct the new file path
        new_file_path = os.path.join(output_directory, f"{new_file_name}.xlsx")

        # Construct the terminal command for file copy
        command = f'copy "{template_path}" "{new_file_path}"'

        # Execute the terminal command
        subprocess.run(command, shell=True, check=True)
        print(f"File duplicated and renamed to: {new_file_path}")
        return new_file_path
    except Exception as e:
        print(f"An error occurred while duplicating and renaming the file: {e}")
        return None

def open_excel_file(file_path):
    try:
        print(f"Opening file: {file_path}")
        # Use subprocess to open the file
        subprocess.run(f'start excel "{file_path}"', shell=True, check=True)
        time.sleep(4)  # Wait for 4 seconds
        print("Excel file opened successfully.")
    except Exception as e:
        print(f"An error occurred while opening the Excel file: {e}")

def process_google_sheet():
    """
    Extract data from a Google Sheet in the next tab and process it.
    """
    try:
        # Switch to the next tab
        pyautogui.hotkey("ctrl", "tab")
        time.sleep(0.5)  # Wait for the tab to load
        pyautogui.hotkey("ctrl", "c")  # Copy selected data
        time.sleep(0.3)
        if __name__ == "__main__":
            # Define paths
            template_file = r"C:\Users\mukun\Downloads\site data\temp.xlsx"  # Update with the actual path
            output_dir = r"C:\Users\mukun\Downloads\site data"  # Directory for the new files

            # Get the data from the clipboard
            clipboard_data = pyperclip.paste().strip()

            if clipboard_data:
                # Create a valid file name by replacing invalid characters and truncating length
                new_file_name = clipboard_data.replace(" ", "_")[:50]  # Replace spaces, limit to 50 characters

                # Duplicate and rename the file
                new_file_path = duplicate_and_rename_file_terminal(template_file, new_file_name, output_dir)

                # Open the file in Excel
                if new_file_path:
                    open_excel_file(new_file_path)
            else:
                print("No data in clipboard to use for file naming.")

        

        

    except Exception as e:
        print(f"An error occurred while processing the Google Sheet: {e}")
        return []
    
def paste_to_first_empty_cell_in_A():
    """
    Navigate to Excel, find the first empty cell in column 'A', and paste the copied data.
    """
    try:
        # Locate and click the Excel icon
        locate_and_click(excel_img)
        time.sleep(0.5)
        
        # Ensure we're in column 'A'
        pyautogui.hotkey("ctrl", "home")  # Move to the top-left corner of the sheet
        time.sleep(0.3)
        pyautogui.press("down")  # Move to the first data row
        
        # Iterate through column 'A' to find the first empty cell
        while True:
            # Copy the current cell's content to clipboard
            pyautogui.hotkey("ctrl", "c")
            time.sleep(0.3)
            cell_data = pyperclip.paste()
            
            if not cell_data.strip():  # If the cell is empty, stop the loop
                break
            
            pyautogui.press("down")  # Move to the next cell
        
        # Paste data into the found cell
        pyautogui.hotkey("ctrl", "v")
        print("Pasted data into the first empty cell in column 'A'.")
        
    except Exception as e:
        print(f"An error occurred while pasting data: {e}")

def locate_and_click(img_path, confidence=0.8, region=None):
    """
    Locate an image on the screen and click its center if found.
    """
    try:
        # Locate the image
        center = pyautogui.locateCenterOnScreen(img_path, region=region, confidence=confidence)
        if center:
            print(f"Image '{img_path}' found at: {center}")
            pyautogui.click(center)  # Click on the center of the located image
        else:
            raise pyautogui.ImageNotFoundException
    except pyautogui.ImageNotFoundException:
        print(f"Image '{img_path}' not found in the specified region.")

def clean_clipboard_content():
    """
    Cleans and formats the clipboard content if it has 3 or fewer rows.
    """
    try:
        # Retrieve clipboard data
        clipboard_data = pyperclip.paste().strip()

        # Split the data into rows
        data_rows = clipboard_data.splitlines()

        # Check if the number of rows is 3 or fewer
        if len(data_rows) <= 3:
            # Clean and process the data
            cleaned_data = "\n".join([line.strip() for line in data_rows])
            
            # Update the clipboard with the cleaned data
            pyperclip.copy(cleaned_data)

            print("Clipboard content cleaned and updated:")
            print(cleaned_data)
        else:
            print(f"Data has {len(data_rows)} rows, which is more than 3. No processing done.")

    except Exception as e:
        print(f"An error occurred: {e}")


def click_and_drag(pos, hold_time=2.0):
    pyautogui.moveTo(pos[0], pos[1], duration=0.3)
    pyautogui.click()
    time.sleep(0.5)
    checkpage("list.png",pos[0], pos[1])
    try:
        # Move to the start position and press the mouse button
        pyautogui.moveTo(23, 843, duration=0.5)
        pyautogui.mouseDown()

        pyautogui.moveTo(1832, 1003, duration=0.4)

        # Hold the mouse button for the specified time
        time.sleep(hold_time)
        print(f"Held for {hold_time} seconds")
        pyautogui.moveTo(1865, 953, duration=0.2)
        time.sleep(0.3)
        # Release the mouse button
        pyautogui.mouseUp()
        print("Mouse released")
        
        # Check for 'corner.png' in the bottom-right 500x500 region
        

        # Copy text to the clipboard
        pyautogui.hotkey("ctrl", "c")
        time.sleep(0.3)
        clean_clipboard_content()
    except Exception as e:
        print(f"An error occurred: {e}")


def checksheet(img_path, region=None):
        try:
            # Locate the image
            center = pyautogui.locateCenterOnScreen(img_path, region=region, confidence=0.8)
            time.sleep(0.5)
            if center:
                print(f"Image '{img_path}' found at: {center}")
                pyautogui.click(center)  # Click on the center of the located image
            else:
                checksheet(img_path)
                raise pyautogui.ImageNotFoundException

        except pyautogui.ImageNotFoundException:
            print(f"Image '{img_path}' not found in the specified region.")
def checkpage(img_path,altimg = None,x=0,y=0):
        try:
            # Locate the image
            center = pyautogui.locateCenterOnScreen(img_path, confidence=0.8)
            time.sleep(0.3)
            if center:
                    
                if x == 0 and y == 0:
                    print(f"Image '{img_path}' found at: {center}")
                else:
                    pyautogui.moveTo(x,y,duration=0.1)
                    pyautogui.click(x,y)  # Click on the center of the located image
                if altimg != None:
                    locate_and_click(altimg, confidence=0.8)
                time.sleep(0.5)
            else:
                raise pyautogui.ImageNotFoundException

        except pyautogui.ImageNotFoundException:
            print(f"Image '{img_path}' not found in the specified region.")



# ----------------------------- Main Execution ----------------------------------------

for j in range(789):
    pyautogui.moveTo(651, 1058, duration=0.3)
    pyautogui.click()
    time.sleep(0.6)
    process_google_sheet()
    pyautogui.moveTo(651, 1058, duration=0.3)
    pyautogui.click()
    
    time.sleep(0.3)
    pyautogui.press("right")
    time.sleep(0.1)
    pyautogui.hotkey("ctrl", "c")
    # Switch back to the original tab
    time.sleep(0.3)
    pyautogui.press("down")
    time.sleep(0.3)
    pyautogui.press("left")
    time.sleep(0.2)
    pyautogui.hotkey("ctrl", "shift", "tab")
        
    locate_and_click(chrome_img, region=taskbar_region)
    time.sleep(0.4)
    pyautogui.hotkey("ctrl", "l")
    time.sleep(0.2)

    pyautogui.hotkey("ctrl", "v")
    time.sleep(0.2)
    pyautogui.press("enter")
    time.sleep(0.8)
    checkpage(lockpage,altimg="continue.png")

    paste_to_first_empty_cell_in_A()
    # Locate and click the chrome icon in the taskbar
    locate_and_click(chrome_img, region=taskbar_region)

    time.sleep(0.4)


    
    # Step 2: Drag from (23, 843) to (1728, 955), hold for 2 seconds, and copy text
    position = (1772, 809)
    click_and_drag(position, hold_time=2)

    # Path to images


    # Step 3: Move to 'excel.png', wait for 0.5 seconds, and paste the data


    # Step 4: Switch to Chrome using Alt + Tab
    pyautogui.hotkey("alt", "tab")
    time.sleep(0.5)
    pyautogui.hotkey("ctrl", "v")

    # Step 5: Check if 'logo.png' is in the top bar of Chrome and click on it
    locate_and_click(chrome_img, region=taskbar_region)
    pyautogui.moveTo(1900, 746, duration=0.4)
    pyautogui.click()


    position = (1764, 834)
    click_and_drag(position, hold_time=2)

    pyautogui.hotkey("alt", "tab")
    time.sleep(0.5)
    pyautogui.press("right")
    time.sleep(0.1)
    pyautogui.press("right")
    time.sleep(0.1)
    pyautogui.press("right")
    time.sleep(0.1)

    pyautogui.hotkey("ctrl", "v")
    time.sleep(0.4)
    pyautogui.press("right")

    # Step 5: Check if 'logo.png' is in the top bar of Chrome and click on it
    locate_and_click(chrome_img, region=taskbar_region)
    pyautogui.moveTo(1900, 746, duration=0.4)
    pyautogui.click()
    
    def sheetmove(imglink = "a"):
        locate_and_click(excel_img, region=taskbar_region)
        time.sleep(0.3)
        if imglink == "a":
            print("EMPTY LINK")
        else:
            pyautogui.hotkey("ctrl","v")
        time.sleep(0.3)
        pyautogui.press("down")
        time.sleep(0.2)
        pyautogui.moveTo(651, 1058, duration=0.3)
        pyautogui.click()
        time.sleep(0.6)
        checksheet("sheetson.png", region=taskbar_region)
        pyautogui.press("enter")
        time.sleep(0.4)
        pyautogui.press("enter")
    position = (1771, 864)
    click_and_drag(position, hold_time=2)

    sheet_img = "sheet.png"
    # Locate and click the chrome icon in the taskbar
    pyautogui.moveTo(651, 1058, duration=0.3)
    pyautogui.click()

    time.sleep(0.3)
    pyautogui.hotkey("ctrl","home") 
    time.sleep(0.2)
    pyautogui.hotkey("ctrl","a")
    time.sleep(0.2)
    pyautogui.hotkey("backspace")
    time.sleep(0.2)
    pyautogui.hotkey("ctrl","v")
    time.sleep(0.2)
    pyautogui.press("right")
    time.sleep(0.2)
    pyautogui.hotkey("ctrl","shift","down")
    time.sleep(0.2)
    pyautogui.hotkey("ctrl", "c")
    # Get data from clipboard and strip extra spaces
    clipboard_data = pyperclip.paste().strip()

            # Split the data into a list of links
    links = clipboard_data.splitlines()

            # Count the number of links
    link_length = len(links)
    print(f"Number of links in clipboard: {link_length}")

    time.sleep(0.2)
    pyautogui.press("down")
    time.sleep(0.2)
    pyautogui.press("up")
    for i in range(link_length):
        time.sleep(0.2)
        pyautogui.hotkey("ctrl", "c")
        locate_and_click(chrome_img, region=taskbar_region)
        time.sleep(0.2)
        pyautogui.hotkey("ctrl", "l")
        time.sleep(0.2)

        pyautogui.hotkey("ctrl", "v")
        time.sleep(0.2)
        pyautogui.press("enter")
        time.sleep(0.8)
        
        checkpage(lockpage,altimg="continue.png")
        pyautogui.moveTo(1900, 746, duration=0.4)
        pyautogui.click()
        time.sleep(0.2)

        
        # Step 2: Drag from (23, 843) to (1728, 955), hold for 2 seconds, and copy text
        position = (1753, 885)
        click_and_drag(position, hold_time=2)
        time.sleep(0.3)
        cell_data = pyperclip.paste().strip()
                
            # Continue if the cell contains "o Data Extracted"
        try:
            # Retrieve the clipboard content and strip any extra spaces

            # Skip if the data is 'o Data Extracted'
            if cell_data != "o Data Extracted":
                # Split the content by new lines
                links = cell_data.splitlines()

                # Truncate to the first four links if there are more than four
                if len(links) > 4:
                    print(f"More than 4 links detected. Truncating to the first 4 links.")
                    links = links[:4]

                # Join the truncated or original links with tabs (\t)
                formatted_data = "\t".join(links)

                print(f"Formatted data: {formatted_data}")
                
                # Update the clipboard with the formatted data
                pyperclip.copy(formatted_data)
                print("Formatted data copied back to clipboard.")
                
                # Call the sheetmove function with the formatted data
                sheetmove(formatted_data)
            else:
                print("Cell contains 'o Data Extracted', skipping...")
                sheetmove("a")
                continue  # Proceed without changing the data in this case
        except Exception as e:
            print(f"An error occurred while processing clipboard data: {e}")

        
    locate_and_click(chrome_img, region=taskbar_region)
    time.sleep(0.2)
    pyautogui.moveTo(1900, 746, duration=0.4)
    pyautogui.click()
    time.sleep(0.2)
    
    locate_and_click(excel_img, region=taskbar_region)
    time.sleep(0.3)
    
    pyautogui.hotkey("ctrl", "s")
    time.sleep(0.4)
    pyautogui.hotkey("ctrl", "w")
    time.sleep(0.4)

        # Locate and click the chrome icon in the taskbar
        
        