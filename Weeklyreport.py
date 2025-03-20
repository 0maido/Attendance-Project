# Developed by Ahmed Abduljalil  aka github username 0maido
# Project name : Weekly Report
# Date : 2025-2-28



import pandas as pd
import tkinter as tk
from tkinter import filedialog, ttk, messagebox
import traceback

# Dictionary to store attendance data
attendance_data = {day: [] for day in ["Monday", "Tuesday", "Wednesday", "Thursday"]}
days_order = ["Monday", "Tuesday", "Wednesday", "Thursday"]

def load_excel(day):
    """Load and process multiple attendance files for a specific day"""
    try:
        file_paths = filedialog.askopenfilenames(
            title=f"Select {day} Attendance File(s)",
            filetypes=[("Excel files", "*.xlsx;*.xls")]
        )
        
        if not file_paths:
            messagebox.showinfo("Cancelled", "No files selected")
            return

        new_data = []
        for file_path in file_paths:
            # Read Excel file
            try:
                df = pd.read_excel(file_path, header=1, engine='openpyxl')
            except:
                df = pd.read_excel(file_path, header=1, engine='xlrd')

            # Validate structure
            if len(df.columns) < 7:
                raise ValueError(f"Invalid file format in {file_path}")

            # Rename columns
            df = df.rename(columns={
                df.columns[2]: "Student Name",
                df.columns[5]: "Student ID"
            }).astype({"Student ID": str})

            # Calculate absences and leaves
            attendance_cols = df.columns[6:]
            df[f"{day}_Absences"] = (df[attendance_cols] == 'A').sum(axis=1)
            df[f"{day}_Leaves"] = (df[attendance_cols] == 'L').sum(axis=1)
            
            new_data.append(df[["Student Name", "Student ID", 
                              f"{day}_Absences", f"{day}_Leaves"]])

        attendance_data[day].extend(new_data)
        status_label.config(text=f"{day}: Loaded {len(new_data)} file(s)")
        update_table()

    except Exception as e:
        error_msg = f"Error loading {day}:\n{str(e)}\n\n{traceback.format_exc()}"
        messagebox.showerror("Loading Error", error_msg)

def aggregate_data():
    """Combine and sum data from all files"""
    all_days = []
    for day in days_order:
        if attendance_data[day]:
            combined = pd.concat(attendance_data[day])
            # Group and sum both absences and leaves
            grouped = combined.groupby(["Student Name", "Student ID"]).agg({
                f"{day}_Absences": 'sum',
                f"{day}_Leaves": 'sum'
            }).reset_index()
            all_days.append(grouped)
    
    if not all_days:
        return pd.DataFrame()
    
    # Merge all days
    merged = all_days[0]
    for df in all_days[1:]:
        merged = merged.merge(df, on=["Student Name", "Student ID"], how='outer')
    
    # Ensure all expected columns exist
    expected_cols = []
    for day in days_order:
        expected_cols.extend([f"{day}_Absences", f"{day}_Leaves"])
    
    for col in expected_cols:
        if col not in merged.columns:
            merged[col] = 0
    
    # Fill missing values and convert to integers
    merged = merged.fillna(0)
    for col in expected_cols:
        merged[col] = merged[col].astype(int)
    
    # Calculate totals
    absence_cols = [f"{day}_Absences" for day in days_order]
    leave_cols = [f"{day}_Leaves" for day in days_order]
    
    merged["Total Absences"] = merged[absence_cols].sum(axis=1)
    merged["Total Leaves"] = merged[leave_cols].sum(axis=1)
    
    return merged

def update_table():
    """Update the GUI table"""
    try:
        tree.delete(*tree.get_children())
        
        merged = aggregate_data()
        if merged.empty:
            status_label.config(text="No data loaded")
            return

        # Strict filtering: only students with absences OR leaves
        filtered = merged[(merged["Total Absences"] > 0) | (merged["Total Leaves"] > 0)]
        
        if filtered.empty:
            status_label.config(text="No absences or leaves recorded")
            return
            
        for _, row in filtered.iterrows():
            values = [row["Student Name"], row["Student ID"]]
            # Add daily absences and leaves
            for day in days_order:
                absences = row[f"{day}_Absences"]
                leaves = row[f"{day}_Leaves"]
                values.append(f"A: {absences}\nL: {leaves}")
            values += [row["Total Absences"], row["Total Leaves"]]
            tree.insert("", "end", values=values)
            
        status_label.config(text=f"Showing {len(filtered)} records")

    except Exception as e:
        messagebox.showerror("Processing Error", f"Table update failed:\n{str(e)}")

def export_to_excel():
    """Export data with formatting"""
    try:
        merged = aggregate_data()
        if merged.empty:
            messagebox.showwarning("No Data", "Load data first!")
            return

        # Create final dataframe with strict filtering
        final_df = merged[(merged["Total Absences"] > 0) | (merged["Total Leaves"] > 0)]
        final_df = final_df[["Student Name", "Student ID"]].copy()
        
        for day in days_order:
            final_df[f"{day} Absences"] = merged[f"{day}_Absences"]
            final_df[f"{day} Leaves"] = merged[f"{day}_Leaves"]
            
        final_df["Total Absences"] = merged["Total Absences"]
        final_df["Total Leaves"] = merged["Total Leaves"]

        # Save file
        file_path = filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            filetypes=[("Excel files", "*.xlsx")]
        )
        if not file_path:
            return

        with pd.ExcelWriter(file_path, engine='xlsxwriter') as writer:
            final_df.to_excel(writer, index=False, sheet_name="Attendance")
            
            workbook = writer.book
            worksheet = writer.sheets["Attendance"]
            
            # Create header formats
            header_red = workbook.add_format({
                'bg_color': '#FFC7CE',
                'bold': True,
                'align': 'center',
                'valign': 'vcenter',
                'border': 1
            })
            
            header_blue = workbook.add_format({
                'bg_color': '#C6EFCE',
                'bold': True,
                'align': 'center',
                'valign': 'vcenter',
                'border': 1
            })
            
            # Apply header formatting
            for col_num, header in enumerate(final_df.columns):
                if "Absences" in header:
                    worksheet.write(0, col_num, header, header_red)
                elif "Leaves" in header:
                    worksheet.write(0, col_num, header, header_blue)
                else:
                    worksheet.write(0, col_num, header)
            
            # Data formatting
            data_format = workbook.add_format({
                'align': 'center',
                'valign': 'vcenter',
                'num_format': '@',
                'border': 1
            })
            
            # Apply data formatting
            for col_num in range(len(final_df.columns)):
                worksheet.set_column(col_num, col_num, 20, data_format)
                
            # Set row heights
            for row_num in range(len(final_df)+1):
                worksheet.set_row(row_num, 25)

        messagebox.showinfo("Success", f"Report exported to:\n{file_path}")

    except Exception as e:
        messagebox.showerror("Export Error", f"Export failed:\n{str(e)}")

# GUI Setup
root = tk.Tk()
root.title("Attendance Manager")
root.geometry("1600x800")

# Configure styles
style = ttk.Style()
style.configure("Treeview", rowheight=40, font=('Arial', 10))
style.configure("Treeview.Heading", font=('Arial', 11, 'bold'))

# Create columns
columns = ["Student Name", "Student ID"]
for day in days_order:
    columns.append(f"{day}")
columns += ["Total Absences", "Total Leaves"]

# Treeview widget
tree = ttk.Treeview(root, columns=columns, show="headings")
for col in columns:
    tree.heading(col, text=col)
    tree.column(col, width=150, anchor='center')

tree.pack(expand=True, fill='both', padx=10, pady=10)

# Control buttons
button_frame = tk.Frame(root)
button_frame.pack(pady=10)

for day in days_order:
    btn = tk.Button(
        button_frame,
        text=f"Add {day}",
        command=lambda d=day: load_excel(d),
        width=12,
        height=2
    )
    btn.pack(side='left', padx=5)

tk.Button(
    button_frame,
    text="Export Report",
    command=export_to_excel,
    width=12,
    height=2
).pack(side='left', padx=5)

# Status label
status_label = tk.Label(root, text="Ready", bd=1, relief=tk.SUNKEN, anchor=tk.W)
status_label.pack(fill='x', padx=10, pady=5)

root.mainloop()
