import tkinter as tk
from openpyxl import Workbook
from tkinter import messagebox
import requests
import webbrowser

def save_to_excel():
    first_name = entry_first_name.get()
    last_name = entry_last_name.get()
    email = entry_email.get()

    # Create an Excel workbook and select the active sheet
    workbook = Workbook()
    sheet = workbook.active

    # Append the data to the Excel sheet
    sheet.append(["First Name", "Last Name", "Email"])
    sheet.append([first_name, last_name, email])

    # Save the workbook to a file
    workbook.save("registration_responses.xlsx")

    # Show a success message box
    messagebox.showinfo("Success", "Registration data saved to Excel.")

def fetch_github_projects():
    github_username = entry_github_username.get()
    api_url = f'https://api.github.com/users/{github_username}/repos'

    try:
        response = requests.get(api_url)
        if response.status_code == 200:
            github_projects = response.json()
            project_listbox.delete(0, tk.END)  # Clear previous entries

            for project in github_projects:
                project_name = project["name"]
                project_listbox.insert(tk.END, project_name)

        else:
            messagebox.showerror("Error", f"Failed to fetch GitHub projects. Status code: {response.status_code}")

    except requests.exceptions.RequestException as e:
        messagebox.showerror("Error", f"Failed to fetch GitHub projects: {e}")

def open_github_project():
    selected_index = project_listbox.curselection()
    if not selected_index:
        return

    selected_project = project_listbox.get(selected_index)
    github_url = f'https://github.com/{entry_github_username.get()}/{selected_project}'
    webbrowser.open_new(github_url)

# Create the main application window
app = tk.Tk()
app.title("Projects Registration Form")

# Create and place widgets on the window
label_first_name = tk.Label(app, text="First Name:")
label_first_name.grid(row=0, column=0)
entry_first_name = tk.Entry(app)
entry_first_name.grid(row=0, column=1)

label_last_name = tk.Label(app, text="Last Name:")
label_last_name.grid(row=1, column=0)
entry_last_name = tk.Entry(app)
entry_last_name.grid(row=1, column=1)

label_email = tk.Label(app, text="Email:")
label_email.grid(row=2, column=0)
entry_email = tk.Entry(app)
entry_email.grid(row=2, column=1)

label_github_username = tk.Label(app, text="GitHub Username:")
label_github_username.grid(row=3, column=0)
entry_github_username = tk.Entry(app)
entry_github_username.grid(row=3, column=1)

submit_button = tk.Button(app, text="Submit", command=save_to_excel)
submit_button.grid(row=4, column=0, columnspan=2)

label_github_projects = tk.Label(app, text="GitHub Projects:")
label_github_projects.grid(row=5, column=0, columnspan=2)
project_listbox = tk.Listbox(app, selectmode=tk.SINGLE)
project_listbox.grid(row=6, column=0, columnspan=2)

fetch_button = tk.Button(app, text="Fetch GitHub Projects", command=fetch_github_projects)
fetch_button.grid(row=7, column=0, columnspan=2)

# Add a button to open the selected GitHub project
open_project_button = tk.Button(app, text="Open Project", command=open_github_project)
open_project_button.grid(row=8, column=0, columnspan=2)

# Include the additional label for more projects registration
label_more_projects = tk.Label(app, text="For more such projects, register yourself.")
label_more_projects.grid(row=9, column=0, columnspan=2)

# Social media links and symbols
label_social_media = tk.Label(app, text="Connect with me:")
label_social_media.grid(row=10, column=0, columnspan=2)

# Start the application's event loop
app.mainloop()
