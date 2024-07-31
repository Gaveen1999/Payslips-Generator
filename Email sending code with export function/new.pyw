import tkinter as tk
from tkinter import filedialog, messagebox, scrolledtext
import os
import pandas as pd
import PyPDF2
import win32com.client as win32

def browse_excel():
    excel_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx")])
    entry_excel_path.delete(0, tk.END)
    entry_excel_path.insert(0, excel_path)

def browse_pdf_folder():
    pdf_folder = filedialog.askdirectory()
    entry_pdf_folder.delete(0, tk.END)
    entry_pdf_folder.insert(0, pdf_folder)

def browse_save_location():
    save_location = filedialog.askdirectory()
    entry_save_location.delete(0, tk.END)
    entry_save_location.insert(0, save_location)

def show_loading_window():
    loading_window = tk.Toplevel()
    loading_window.title("Sending Emails")
    loading_label = tk.Label(loading_window, text="Sending emails, please wait...")
    loading_label.pack(padx=20, pady=20)
    loading_window.grab_set()
    root.update()  # Update the main window to display the loading window
    return loading_window

def encrypt_and_send_emails():
    excel_path = entry_excel_path.get()
    pdf_folder = entry_pdf_folder.get()
    save_location = entry_save_location.get()
    on_behalf_of = entry_on_behalf_of.get()  # Get the on behalf of email address

    if not os.path.isfile(excel_path):
        messagebox.showerror("Error", "Please select a valid Excel file.")
        return

    if not os.path.isdir(pdf_folder):
        messagebox.showerror("Error", "Please select a valid PDF folder.")
        return

    if not os.path.isdir(save_location):
        messagebox.showerror("Error", "Please select a valid save location.")
        return

    df = pd.read_excel(excel_path)
    pdf_files = [file for file in os.listdir(pdf_folder) if file.endswith('.pdf')]

    outlook = win32.Dispatch('Outlook.Application')

    loading_window = show_loading_window()  # Show loading window

    success_logs = {}
    error_logs = {}

    total_success_count = 0
    total_error_count = 0

    for index, row in df.iterrows():
        user_id = str(row['user ID'])
        birthday = str(row['birthday'])
        email = row['email']
        username = row['username']  # Assuming 'username' is the column name in Excel containing the username

        filename = user_id + '.pdf'
        if filename in pdf_files:
            pdf_path = os.path.join(pdf_folder, filename)

            pdf_writer = PyPDF2.PdfWriter()
            pdf_reader = PyPDF2.PdfReader(pdf_path)
            for page_num in range(len(pdf_reader.pages)):
                pdf_writer.add_page(pdf_reader.pages[page_num])
            pdf_writer.encrypt(user_id + birthday)

            encrypted_pdf_path = os.path.join(save_location, f'encrypted_{filename}')
            with open(encrypted_pdf_path, 'wb') as encrypted_file:
                pdf_writer.write(encrypted_file)

            mail = outlook.CreateItem(0)
            mail.Subject = f"{username} | Salary e-statement"
            mail.To = email
            mail.Attachments.Add(Source=encrypted_pdf_path, Type=17)
            mail.HTMLBody = f"""\
{username},

<ol>
  <li>The detailed Salary e-statement for TKSS is attached herewith.</li>
  <li>Kindly insert 4 digits of the employee ID followed by DOB(DDMM-Format) to open the attachment.</li>
  <li>Eg: If Employee ID is 8004 and the DOB is 14071981, password should be 80041407.(DOB - Date of Birth).</li>
</ol>

<b>Thanks and Regards,</b></br>
<b>HR Department.</b></br>
<img src = "1.png">
"""
            # Set the on behalf of email address
            mail.SentOnBehalfOfName = on_behalf_of

            try:
                mail.Send()
                success_logs[user_id] = success_logs.get(user_id, 0) + 1
                total_success_count += 1
            except Exception as e:
                error_logs[user_id] = error_logs.get(user_id, 0) + 1
                total_error_count += 1
                continue
        else:
            error_logs[user_id] = error_logs.get(user_id, 0) + 1
            total_error_count += 1

    loading_window.destroy()  # Close loading window after sending emails

    event_viewer_text.delete(1.0, tk.END)
    event_viewer_text.insert(tk.END, f"Total Success Count: {total_success_count}\n")
    event_viewer_text.insert(tk.END, f"Total Error Count: {total_error_count}\n")
    event_viewer_text.insert(tk.END, "\nSuccess Logs per User ID:\n")
    for user_id, count in success_logs.items():
        event_viewer_text.insert(tk.END, f"User ID: {user_id}, User Name : {count}\n")
    event_viewer_text.insert(tk.END, "\nError Logs per User ID:\n")
    for user_id, count in error_logs.items():
        event_viewer_text.insert(tk.END, f"User ID: {user_id}, User Name: {count}\n")


def export_logs_to_text():
    logs_text = event_viewer_text.get("1.0", tk.END)
    export_path = filedialog.asksaveasfilename(defaultextension=".txt", filetypes=[("Text files", "*.txt")])
    if export_path:
        with open(export_path, "w") as file:
            file.write(logs_text)
        messagebox.showinfo("Export Complete", "Event logs exported to text file successfully.")



# Create the main window
root = tk.Tk()
root.title("PDF Encryption and Email Sender - Teknowledge Shared Services")

# Load the company logo
logo_image = tk.PhotoImage(file="1.png")

# Create and place widgets
company_logo_label = tk.Label(root, image=logo_image)
company_logo_label.grid(row=0, column=0, columnspan=3)

label_excel_path = tk.Label(root, text="Excel File:")
label_excel_path.grid(row=1, column=0, padx=5, pady=5)
entry_excel_path = tk.Entry(root, width=50)
entry_excel_path.grid(row=1, column=1, padx=5, pady=5)
button_browse_excel = tk.Button(root, text="Browse", command=browse_excel)
button_browse_excel.grid(row=1, column=2, padx=5, pady=5)

label_pdf_folder = tk.Label(root, text="Generated Payslips Folder:")
label_pdf_folder.grid(row=2, column=0, padx=5, pady=5)
entry_pdf_folder = tk.Entry(root, width=50)
entry_pdf_folder.grid(row=2, column=1, padx=5, pady=5)
button_browse_pdf_folder = tk.Button(root, text="Browse", command=browse_pdf_folder)
button_browse_pdf_folder.grid(row=2, column=2, padx=5, pady=5)

label_save_location = tk.Label(root, text="Save Location:")
label_save_location.grid(row=3, column=0, padx=5, pady=5)
entry_save_location = tk.Entry(root, width=50)
entry_save_location.grid(row=3, column=1, padx=5, pady=5)
button_browse_save_location = tk.Button(root, text="Browse", command=browse_save_location)
button_browse_save_location.grid(row=3, column=2, padx=5, pady=5)

label_on_behalf_of = tk.Label(root, text="On Behalf Of Email:")
label_on_behalf_of.grid(row=4, column=0, padx=5, pady=5)
entry_on_behalf_of = tk.Entry(root, width=50)
entry_on_behalf_of.grid(row=4, column=1, padx=5, pady=5)

button_process = tk.Button(root, text="Encrypt and Send Emails", command=encrypt_and_send_emails)
button_process.grid(row=5, column=1, padx=5, pady=5)

event_viewer_label = tk.Label(root, text="Event Viewer:")
event_viewer_label.grid(row=6, column=0, padx=5, pady=5)
event_viewer_text = scrolledtext.ScrolledText(root, width=50, height=10)
event_viewer_text.grid(row=6, column=1, columnspan=2, padx=5, pady=5)


# Add Export Logs to Text button
button_export_logs_text = tk.Button(root, text="Export Logs", command=export_logs_to_text)
button_export_logs_text.grid(row=7, column=1, padx=5, pady=5)


root.mainloop()