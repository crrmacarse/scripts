import tkinter as tk
from tkinter import filedialog, messagebox

# Import modules
try:
    from scraper import scrape
except ImportError:
    def scrape(url):
        raise ImportError("The scraper module is missing or has errors.")

try:
    from get_coordinates import get_coordinates
except ImportError:
    def get_coordinates(location):
        raise ImportError("The get_coordinates module is missing or has errors.")

# try:
#     from is_weather_good_here import is_weather_good_here
# except ImportError:
#     def is_weather_good_here(location):
#         raise ImportError("The get_coordinates module is missing or has errors.")

def scrape_website():
    url = website_url_entry.get()
    if not url:
        messagebox.showerror("Error", "Please enter a website URL.")
        return
    try:
        result = scrape(url)
        messagebox.showinfo("Success", f"Website scraped successfully!\n\n{result[:500]}...")
    except Exception as e:
        messagebox.showerror("Error", f"An error occurred: {e}")

def get_location_coordinates():
    location = location_entry.get()
    if not location:
        messagebox.showerror("Error", "Please enter a location.")
        return
    try:
        coordinates = get_coordinates(location)
        
        messagebox.showinfo("Success", f"Coordinates:\n {coordinates["lat"]}, {coordinates["lng"]}")
    except Exception as e:
        messagebox.showerror("Error", f"An error occurred: {e}")

# Create the main window
root = tk.Tk()
root.title("Python Script GUI")

# Menu Bar
menu = tk.Menu(root)
root.config(menu=menu)
file_menu = tk.Menu(menu, tearoff=0)
menu.add_cascade(label="File", menu=file_menu)
file_menu.add_command(label="Exit", command=root.quit)

# Web Scraper Section
tk.Label(root, text="Web Scraper").grid(row=2, column=0, pady=10, padx=10, sticky="w")
website_url_entry = tk.Entry(root, width=50)
website_url_entry.grid(row=3, column=0, padx=10, pady=5)
tk.Button(root, text="Scrape Website", command=scrape_website).grid(row=3, column=1, padx=10, pady=5)

# Get Coordinates Section
tk.Label(root, text="Get Location Coordinates").grid(row=4, column=0, pady=10, padx=10, sticky="w")
location_entry = tk.Entry(root, width=50)
location_entry.grid(row=5, column=0, padx=10, pady=5)
tk.Button(root, text="Get Coordinates", command=get_location_coordinates).grid(row=5, column=1, padx=10, pady=5)

# Status Bar
status = tk.Label(root, text="Ready", bd=1, relief=tk.SUNKEN, anchor="w")
status.grid(row=6, column=0, columnspan=2, sticky="we")

# Run the GUI
root.mainloop()