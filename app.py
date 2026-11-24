
import customtkinter as ctk
import json
import os
import uuid
from datetime import datetime
from tkinter import filedialog
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment
from matplotlib.figure import Figure
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
editing_expense = None
visible_expenses = []

# Appearance settings
ctk.set_appearance_mode("dark")
ctk.set_default_color_theme("blue")
def load_expenses():

    if not os.path.exists("expenses.json"):
        return []

    try:
        with open("expenses.json", "r", encoding="utf-8") as file:
            expenses = json.load(file)

        updated = False

        for expense in expenses:
            if "id" not in expense:
                expense["id"] = str(uuid.uuid4())
                updated = True

        if updated:
            with open("expenses.json", "w", encoding="utf-8") as file:
                json.dump(expenses, file, indent=4)

        return expenses

    except (json.JSONDecodeError, OSError):
        return []

   

def calculate_total():

    expenses = load_expenses()

    total = 0

    for expense in expenses:
        total += float(expense.get("amount", 0))

    return total

def count_expenses():
    expenses = load_expenses()
    return len(expenses)


def calculate_average():

    expenses = load_expenses()

    if not expenses:
        return 0

    total = sum(
        float(expense.get("amount", 0))
        for expense in expenses
    )

    return total / len(expenses)

def count_categories():
    expenses = load_expenses()

    categories = set()

    for expense in expenses:
        categories.add(
            expense.get("category", "Other")
        )

    return len(categories)


def save_expense():
    name = name_entry.get().strip()
    category = category_menu.get()
    amount_text = amount_entry.get().strip()
    date = date_entry.get().strip()

    if not name or not amount_text or not date:
        status_label.configure(
            text="Please fill in all fields.",
            text_color="#F87171"
        )
        return
    
    try:
        datetime.strptime(date, "%Y-%m-%d")

    except ValueError:
        status_label.configure(
            text="Please enter a valid date (YYYY-MM-DD).",
            text_color="#F87171"
        )
        return


    try:
        amount = float(amount_text)

        if amount <= 0:
            raise ValueError

    except ValueError:
        status_label.configure(
            text="Please enter a valid positive amount.",
            text_color="#F87171"
        )
        return

    new_expense = {
        "id": str(uuid.uuid4()),
        "name": name,
        "category": category,
        "amount": amount,
        "date": date
    }

    global editing_expense

    was_editing = editing_expense is not None
    expenses = load_expenses()

    if editing_expense is not None:

        expense_found = False

        for index, expense in enumerate(expenses):

            if expense.get("id") == editing_expense:
                new_expense["id"] = editing_expense
                expenses[index] = new_expense
                expense_found = True
                break

        editing_expense = None

        if not expense_found:
            status_label.configure(
                text="Expense not found. Please try again.",
                text_color="#F87171"
            )
            return

    else:
        expenses.append(new_expense)

    try:
        with open("expenses.json", "w", encoding="utf-8") as file:
            json.dump(expenses, file, indent=4)

    except OSError:
        status_label.configure(
        text="Could not save expense. Please try again.",
        text_color="#F87171"
    )
        return

    if was_editing:
        status_label.configure(
            text="Expense updated successfully!",
            text_color="#4ADE80"
        )
    else:
        status_label.configure(
            text="Expense added successfully!",
            text_color="#4ADE80"
        )

    name_entry.delete(0, "end")
    amount_entry.delete(0, "end")
    date_entry.delete(0, "end")

    form_title.configure(text="Add New Expense")
    add_button.configure(text="Add Expense")
    cancel_button.grid_remove()

    refresh_dashboard()



def export_expenses():
    expenses = visible_expenses.copy()

    if not expenses:
        status_label.configure(
            text="No expenses available to export.",
            text_color="#F87171"
        )
        return

    file_path = filedialog.asksaveasfilename(
        defaultextension=".xlsx",
        filetypes=[("Excel Files", "*.xlsx")],
        title="Save Excel File"
    )

    if not file_path:
        return

    try:
        workbook = Workbook()
        sheet = workbook.active
        sheet.title = "Expenses"

        headers = ["Name", "Category", "Amount", "Date"]
        sheet.append(headers)

        for expense in expenses:
            sheet.append([
                expense.get("name", ""),
                expense.get("category", ""),
                expense.get("amount", 0),
                expense.get("date", "")
            ])

        # Style the header
        for cell in sheet[1]:
            cell.font = Font(
                bold=True,
                color="FFFFFF"
            )
            cell.fill = PatternFill(
                start_color="0F766E",
                end_color="0F766E",
                fill_type="solid"
            )
            cell.alignment = Alignment(
                horizontal="center"
            )

        # Adjust column widths
        sheet.column_dimensions["A"].width = 25
        sheet.column_dimensions["B"].width = 20
        sheet.column_dimensions["C"].width = 15
        sheet.column_dimensions["D"].width = 18

        # Enable filters
        sheet.auto_filter.ref = sheet.dimensions

        # Freeze header row
        sheet.freeze_panes = "A2"

        workbook.save(file_path)

        status_label.configure(
            text="Excel file exported successfully!",
            text_color="#4ADE80"
        )

    except OSError:
        status_label.configure(
            text="Could not export Excel file.",
            text_color="#F87171"
        )



def clear_filters():
    search_entry.delete(0, "end")
    category_filter.set("All Categories")
    sort_menu.set("Default")
    refresh_dashboard()

def delete_expense(expense_id):
    expenses = load_expenses()

    expenses = [
        expense for expense in expenses
        if expense.get("id") != expense_id
    ]

    with open("expenses.json", "w", encoding="utf-8") as file:
        json.dump(expenses, file, indent=4)

    refresh_dashboard()    

def edit_expense(expense):
    global editing_expense
    editing_expense = expense["id"]

    name_entry.delete(0, "end")
    name_entry.insert(0, expense.get("name", ""))

    category_menu.set(expense.get("category", "Other"))

    amount_entry.delete(0, "end")
    amount_entry.insert(0, str(expense.get("amount", "")))

    date_entry.delete(0, "end")
    date_entry.insert(0, expense.get("date", ""))

    form_title.configure(text="Edit Expense")
    add_button.configure(text="Update Expense")
    cancel_button.grid()

    status_label.configure(
        text="Editing selected expense...",
        text_color="#38BDF8"
    )      

def cancel_edit():
    global editing_expense

    editing_expense = None

    name_entry.delete(0, "end")
    amount_entry.delete(0, "end")
    date_entry.delete(0, "end")

    category_menu.set("Food")

    form_title.configure(text="Add New Expense")
    add_button.configure(text="Add Expense")
    cancel_button.grid_remove()

    status_label.configure(
        text="Editing cancelled.",
        text_color="#FBBF24"
    )    


def refresh_dashboard():
    global visible_expenses

    # Load expenses once
    all_expenses = load_expenses()

    # Calculate dashboard statistics
    total = sum(
        float(expense.get("amount", 0))
        for expense in all_expenses
    )

    transaction_count = len(all_expenses)

    average = (
        total / transaction_count
        if transaction_count > 0
        else 0
    )

    categories = {
        expense.get("category", "Other")
        for expense in all_expenses
    }

    updated_values = [
        f"${total:,.2f}",
        str(transaction_count),
        f"${average:,.2f}",
        str(len(categories))
    ]

    for index, value in enumerate(updated_values):
        stat_value_labels[index].configure(text=value)
    refresh_reports()
    refresh_categories()

    # Remove old transaction rows
    for widget in table_frame.winfo_children():
        if int(widget.grid_info()["row"]) > 0:
            widget.destroy()

    # Load updated transactions
    expenses = all_expenses.copy()

    # Search and Category Filter
    search_text = search_entry.get().strip().lower()
    selected_category = category_filter.get()

    filtered_expenses = []

    for expense in expenses:
        name_matches = (
            search_text in expense.get("name", "").lower()
        )

        category_matches = (
            selected_category == "All Categories"
            or expense.get("category", "Other") == selected_category
        )

        if name_matches and category_matches:
            filtered_expenses.append(expense)

    expenses = filtered_expenses
    # Sort Transactions
    sort_option = sort_menu.get()

    if sort_option == "Highest Amount":
        expenses.sort(
            key=lambda expense: float(expense.get("amount", 0)),
            reverse=True
        )

    elif sort_option == "Lowest Amount":
        expenses.sort(
            key=lambda expense: float(expense.get("amount", 0))
        )

    elif sort_option == "Newest First":
        expenses.sort(
            key=lambda expense: expense.get("date", ""),
            reverse=True
        )

    elif sort_option == "Oldest First":
        expenses.sort(
            key=lambda expense: expense.get("date", "")
        )
    visible_expenses = expenses.copy()
    # Empty Table Message
    if not expenses:
        empty_label = ctk.CTkLabel(
            table_frame,
            text="No transactions found.",
            text_color=("#64748B", "#94A3B8"),
            font=("Arial", 14)
        )
        empty_label.grid(
            row=1,
            column=0,
            columnspan=6,
            pady=20
        )
        
        return

    # Display updated transactions
    for row_index, expense in enumerate(expenses, start=1):
        values = [
            expense.get("name", "Unknown"),
            expense.get("category", "Other"),
            f"${float(expense.get('amount', 0)):,.2f}",
            expense.get("date", "Unknown")
        ]

        for column_index, value in enumerate(values):
            cell = ctk.CTkLabel(
                table_frame,
                text=str(value),
                font=("Arial", 13),
                text_color=("#1E293B", "#F8FAFC")
            )

            cell.grid(
                row=row_index,
                column=column_index,
                padx=12,
                pady=12,
                sticky="w"
            )

        delete_button = ctk.CTkButton(
            table_frame,
            text="Delete",
            width=70,
            height=28,
            fg_color="#DC2626",
            hover_color="#B91C1C",
            command=lambda expense_id=expense["id"]: delete_expense(
                expense_id
))
        

        delete_button.grid(
            row=row_index,
            column=4,
            padx=12,
            pady=8
        )
        edit_button = ctk.CTkButton(
            table_frame,
            text="Edit",
            width=70,
            height=28,
            fg_color="#2563EB",
            hover_color="#1D4ED8",
            command=lambda expense=expense: edit_expense(expense)
        )

        edit_button.grid(
            row=row_index,
            column=5,
            padx=12,
            pady=8
        )

# Main window
app = ctk.CTk()

app.title("Expense Tracker")
app.geometry("1100x700")
app.minsize(900, 600)

app.configure(fg_color=("#F1F5F9", "#0F172A"))

# Main layout
app.grid_columnconfigure(1, weight=1)
app.grid_rowconfigure(0, weight=1)

# Sidebar
sidebar = ctk.CTkFrame(
    app,
    width=220,
    corner_radius=0,
    fg_color=("#E2E8F0", "#111827")
)

sidebar.grid(
    row=0,
    column=0,
    sticky="nsew"
)

sidebar.grid_propagate(False)

# Sidebar title
logo = ctk.CTkLabel(
    sidebar,
    text="EXPENSE TRACKER",
    font=("Arial", 20, "bold"),
    text_color="#14B8A6"
)

logo.pack(pady=(35, 45), padx=15)

def go_to_dashboard():
    content.update_idletasks()
    content._parent_canvas.yview_moveto(0)

def go_to_transactions():
    content.update_idletasks()
    content._parent_canvas.yview_moveto(0.10)    

def go_to_categories():
    content.update_idletasks()
    content._parent_canvas.yview_moveto(0.25)

def go_to_reports():
    content.update_idletasks()
    content._parent_canvas.yview_moveto(1.0)   

def change_appearance(mode):

    ctk.set_appearance_mode(mode)
    app.after(250, refresh_reports)

    app.after(150, app.lift)
    app.after(200, app.focus_force)   

def go_to_settings():
    content.update_idletasks()
    content._parent_canvas.yview_moveto(1.0)
    app.lift()
    app.focus_force()

# Sidebar buttons
menu_items = [
    "Dashboard",
    "Transactions",
    "Categories",
    "Reports",
    "Settings"
]

for item in menu_items:

    button = ctk.CTkButton(
        sidebar,
        text=item,
        height=42,
        corner_radius=8,
        fg_color="transparent",
        hover_color=("#CBD5E1", "#263449"),
        text_color=("#0F172A", "#F8FAFC"),
        anchor="w",
        command=(
    go_to_transactions if item == "Transactions"
    else go_to_categories if item == "Categories"
    else go_to_reports if item == "Reports"
    else go_to_settings if item == "Settings"
    else go_to_dashboard
)
    )

    button.pack(
        fill="x",
        padx=15,
        pady=5
    )

# Main content
content = ctk.CTkScrollableFrame(
    app,
    corner_radius=0,
    fg_color=("#F1F5F9", "#0F172A")
)

content.grid(
    row=0,
    column=1,
    sticky="nsew",
    padx=25,
    pady=25
)

content.grid_columnconfigure((0, 1), weight=1)

# Dashboard title
title = ctk.CTkLabel(
    content,
    text="Dashboard Overview",
    font=("Arial", 28, "bold"),
    text_color=("#0F172A", "#F8FAFC")
)

title.grid(
    row=0,
    column=0,
    sticky="w",
    pady=(10, 5)
)

subtitle = ctk.CTkLabel(
    content,
    text="Welcome to your personal expense manager.",
    font=("Arial", 14),
    text_color=("#64748B", "#94A3B8")
)

subtitle.grid(
    row=1,
    column=0,
    sticky="w",
    pady=(0, 30)
)

# First dashboard card

# Dashboard Statistics

cards_data = [
    {
        "title": "Total Expenses",
        "value": f"${calculate_total():,.2f}",
        "color": "#14B8A6",
        "icon": "💰"
    },
    {
        "title": "Transactions",
        "value": str(count_expenses()),
        "color": "#A78BFA",
        "icon": "🔄"
    },
    {
        "title": "Average Expense",
        "value": f"${calculate_average():,.2f}",
        "color": "#38BDF8",
        "icon": "📊"
    },
    {
        "title": "Categories",
        "value": str(count_categories()),
        "color": "#FBBF24",
        "icon": "📁"
    }
]
stat_value_labels = []
for index, data in enumerate(cards_data):

    row = 2 + (index // 2)
    column = index % 2

    card = ctk.CTkFrame(
        content,
        height=140,
        corner_radius=15,
        fg_color=("#FFFFFF", "#1E293B")
    )

    card.grid(
        row=row,
        column=column,
        sticky="ew",
        padx=10,
        pady=10
    )

    card.grid_propagate(False)
    # Card Icon
    icon_label = ctk.CTkLabel(
        card,
        text=data["icon"],
        font=("Segoe UI Emoji", 32)
    )

    icon_label.place(
        x=20,
        rely=0.5,
        anchor="w"
    )

    card_title = ctk.CTkLabel(
        card,
        text=data["title"],
        font=("Arial", 16),
        text_color=("#64748B", "#94A3B8")
    )

    card_title.pack(
        anchor="w",
        padx=85,
        pady=(25, 10)
    )

    card_value = ctk.CTkLabel(
        card,
        text=data["value"],
        font=("Arial", 28, "bold"),
        text_color=data["color"]
    )
    stat_value_labels.append(card_value)
    card_value.pack(
        anchor="w",
        padx=85
    )

# Transactions Section

transactions_title = ctk.CTkLabel(
    content,
    text="Recent Transactions",
    font=("Arial", 22, "bold"),
    text_color=("#0F172A", "#F8FAFC")
)

transactions_title.grid(
    row=4,
    column=0,
    columnspan=2,
    sticky="w",
    padx=10,
    pady=(30, 15)
)
# Search Expenses

# Search and Category Filter Container

filter_frame = ctk.CTkFrame(
    content,
    fg_color="transparent"
)

filter_frame.grid(
    row=5,
    column=0,
    columnspan=2,
    padx=10,
    pady=10,
    sticky="w"
)

# Search Expenses

search_entry = ctk.CTkEntry(
    filter_frame,
    placeholder_text="Search expenses...",
    width=300,
    height=40,
    corner_radius=10
)

search_entry.pack(
    side="left",
    padx=(0, 10)
)

search_entry.bind(
    "<Return>",
    lambda event: refresh_dashboard()
)

# Search Button

search_button = ctk.CTkButton(
    filter_frame,
    text="Search",
    width=100,
    height=40,
    command=refresh_dashboard
)

search_button.pack(
    side="left",
    padx=(0, 10)
)

# Category Filter

category_filter = ctk.CTkOptionMenu(
    filter_frame,
    values=[
       "All Categories",
    "Food",
    "Bills",
    "Transport",
    "Education",
    "Entertainment",
    "Health",
    "Shopping",
    "Other"
],
    command=lambda choice: refresh_dashboard(),
    width=170,
    height=40
)

category_filter.pack(side="left")
# Clear Filters Button

clear_button = ctk.CTkButton(
    filter_frame,
    text="Clear Filters",
    command=clear_filters
)
clear_button.pack(side="left", padx=10)


# Export Button
export_button = ctk.CTkButton(
    filter_frame,
    text="Export to Excel",
    width=140,
    height=40,
    fg_color="#0F766E",
    hover_color="#115E59",
    command=export_expenses
)

export_button.pack(
    side="left",
    padx=(10, 0)
)


# Sort Transactions

sort_menu = ctk.CTkOptionMenu(
    filter_frame,
    values=[
        "Default",
        "Highest Amount",
        "Lowest Amount",
        "Newest First",
        "Oldest First"
    ],
    width=160,
    height=40,
    command=lambda choice: refresh_dashboard()
)

sort_menu.pack(
    side="left",
    padx=(10, 0)
)

# Table Container

table_frame = ctk.CTkFrame(
    content,
    fg_color=("#FFFFFF", "#1E293B"),
    corner_radius=12
)

table_frame.grid(
    row=6,
    column=0,
    columnspan=2,
    sticky="ew",
    padx=10,
    pady=5
)

# Table Columns

columns = ["Name", "Category", "Amount", "Date", "Actions"]

for column in range(6):
    table_frame.grid_columnconfigure(
        column,
        weight=1
    )

# Table Header

for index, column_name in enumerate(columns):

    header = ctk.CTkLabel(
        table_frame,
        text=column_name,
        font=("Arial", 14, "bold"),
        text_color="#14B8A6"
    )

    if column_name == "Actions":
        header.grid(
            row=0,
            column=index,
            columnspan=2,
            padx=12,
            pady=15,
            sticky="w"
        )
    else:
        header.grid(
            row=0,
            column=index,
            padx=12,
            pady=15,
            sticky="w"
        )

# Add Expense Form

form_title = ctk.CTkLabel(
    content,
    text="Add New Expense",
    font=("Arial", 22, "bold"),
    text_color=("#0F172A", "#F8FAFC")
)
form_title.grid(
    row=7,
    column=0,
    columnspan=2,
    padx=25,
    pady=(30, 15),
    sticky="w"
)

form_frame = ctk.CTkFrame(
    content,
    fg_color=("#FFFFFF", "#1E293B"),
    corner_radius=15
)
form_frame.grid(
    row=8,
    column=0,
    columnspan=2,
    padx=25,
    pady=10,
    sticky="ew"
)

# Expense Name
name_label = ctk.CTkLabel(
    form_frame,
    text="Expense Name",
    text_color=("#0F172A", "#F8FAFC")
)
name_label.grid(row=0, column=0, padx=15, pady=10, sticky="w")

name_entry = ctk.CTkEntry(
    form_frame,
    placeholder_text="Enter expense name",
    width=250
)
name_entry.grid(row=1, column=0, padx=15, pady=5)

# Category
category_label = ctk.CTkLabel(
    form_frame,
    text="Category",
    text_color=("#0F172A", "#F8FAFC")
)
category_label.grid(row=0, column=1, padx=15, pady=10, sticky="w")

category_menu = ctk.CTkOptionMenu(
    form_frame,
    values=[
        "Food",
        "Bills",
        "Transport",
        "Education",
        "Entertainment",
        "Health",
        "Shopping",
        "Other"
    ]
)
category_menu.grid(row=1, column=1, padx=15, pady=5)

# Amount
amount_label = ctk.CTkLabel(
    form_frame,
    text="Amount",
    text_color=("#0F172A", "#F8FAFC")
)
amount_label.grid(row=2, column=0, padx=15, pady=10, sticky="w")

amount_entry = ctk.CTkEntry(
    form_frame,
    placeholder_text="Enter amount"
)
amount_entry.grid(row=3, column=0, padx=15, pady=5)

# Date
date_label = ctk.CTkLabel(
    form_frame,
    text="Date (YYYY-MM-DD)",
    text_color=("#0F172A", "#F8FAFC")
)
date_label.grid(row=2, column=1, padx=15, pady=10, sticky="w")

date_entry = ctk.CTkEntry(
    form_frame,
    placeholder_text="2026-11-21"
)
date_entry.grid(row=3, column=1, padx=15, pady=5)

# Status Message
status_label = ctk.CTkLabel(
    form_frame,
    text="",
    font=("Arial", 12)
)
status_label.grid(
    row=4,
    column=0,
    columnspan=2,
    padx=15,
    pady=10
)

# Add Expense Button
add_button = ctk.CTkButton(
    form_frame,
    text="Add Expense",
    command=save_expense,
    width=200,
    height=40,
    corner_radius=10,
    fg_color="#2563EB",
    hover_color="#1D4ED8"
)
add_button.grid(
    row=5,
    column=0,
    columnspan=1,
    padx=15,
    pady=15
)
cancel_button = ctk.CTkButton(
    form_frame,
    text="Cancel Edit",
    command=cancel_edit,
    width=200,
    height=40,
    corner_radius=10,
    fg_color="#475569",
    hover_color="#334155"
)

cancel_button.grid(
    row=5,
    column=1,
    padx=15,
    pady=15
)

cancel_button.grid_remove()
# Categories Section

categories_title = ctk.CTkLabel(
    content,
    text="Expense Categories",
    font=("Arial", 26, "bold"),
    text_color=("#0F172A", "#F8FAFC")
)

categories_title.grid(
    row=9,
    column=0,
    columnspan=2,
    padx=25,
    pady=(35, 15),
    sticky="w"
)
categories_frame = ctk.CTkFrame(
    content,
    fg_color=("#FFFFFF", "#1E293B"),
    corner_radius=15
)

categories_frame.grid(
    row=10,
    column=0,
    columnspan=2,
    padx=10,
    pady=10,
    sticky="ew"
)
categories_frame.grid_columnconfigure(
    (0, 1, 2),
    weight=1
)
category_headers = [
    "Category",
    "Total Expenses",
    "Transactions"
]

for index, header_text in enumerate(category_headers):

    header = ctk.CTkLabel(
        categories_frame,
        text=header_text,
        font=("Arial", 14, "bold"),
        text_color="#14B8A6"
    )

    header.grid(
        row=0,
        column=index,
        padx=15,
        pady=15,
        sticky="w"
    )
# Reports Section

reports_title = ctk.CTkLabel(
    content,
    text="Financial Reports",
    font=("Arial", 26, "bold"),
    text_color=("#0F172A", "#F8FAFC")
)

reports_title.grid(
    row=11,
    column=0,
    columnspan=2,
    padx=25,
    pady=(35, 15),
    sticky="w"
)
# Reports - Total Expenses Card

report_total_card = ctk.CTkFrame(
    content,
    fg_color=("#FFFFFF", "#1E293B"),
    corner_radius=15,
    height=130
)

report_total_card.grid(
    row=12,
    column=0,
    sticky="ew",
    padx=10,
    pady=10
)

report_total_card.grid_propagate(False)

report_total_title = ctk.CTkLabel(
    report_total_card,
    text="Total Expenses",
    font=("Arial", 16),
    text_color=("#475569", "#94A3B8")
)
# Reports - Highest Expense Card

report_highest_card = ctk.CTkFrame(
    content,
    fg_color=("#FFFFFF", "#1E293B"),
    corner_radius=15,
    height=130
)

report_highest_card.grid(
    row=12,
    column=1,
    sticky="ew",
    padx=10,
    pady=10
)

report_highest_card.grid_propagate(False)

report_highest_title = ctk.CTkLabel(
    report_highest_card,
    text="Highest Expense",
    font=("Arial", 16),
    text_color=("#64748B", "#94A3B8")
)

report_highest_title.pack(
    anchor="w",
    padx=20,
    pady=(20, 5)
)

expenses = load_expenses()

highest_amount = max(
    [float(expense.get("amount", 0)) for expense in expenses],
    default=0
)

report_highest_value = ctk.CTkLabel(
    report_highest_card,
    text=f"${highest_amount:,.2f}",
    font=("Arial", 26, "bold"),
    text_color="#A78BFA"
)

report_highest_value.pack(
    anchor="w",
    padx=20
)
report_total_title.pack(
    anchor="w",
    padx=20,
    pady=(20, 5)
)

report_total_value = ctk.CTkLabel(
    report_total_card,
    text=f"${calculate_total():,.2f}",
    font=("Arial", 26, "bold"),
    text_color="#14B8A6"
)

report_total_value.pack(
    anchor="w",
    padx=20
)
# Reports - Lowest Expense Card

report_lowest_card = ctk.CTkFrame(
    content,
    fg_color=("#FFFFFF", "#1E293B"),
    corner_radius=15,
    height=130
)

report_lowest_card.grid(
    row=13,
    column=0,
    sticky="ew",
    padx=10,
    pady=10
)

report_lowest_card.grid_propagate(False)

report_lowest_title = ctk.CTkLabel(
    report_lowest_card,
    text="Lowest Expense",
    font=("Arial", 16),
    text_color=("#64748B", "#94A3B8")
)

report_lowest_title.pack(
    anchor="w",
    padx=20,
    pady=(20, 5)
)
# Reports - Total Categories Card

report_categories_card = ctk.CTkFrame(
    content,
    fg_color=("#FFFFFF", "#1E293B"),
    corner_radius=15,
    height=130
)

report_categories_card.grid(
    row=13,
    column=1,
    sticky="ew",
    padx=10,
    pady=10
)

report_categories_card.grid_propagate(False)

report_categories_title = ctk.CTkLabel(
    report_categories_card,
    text="Total Categories",
    font=("Arial", 16),
    text_color=("#64748B", "#94A3B8")
)

report_categories_title.pack(
    anchor="w",
    padx=20,
    pady=(20, 5)
)

expenses = load_expenses()

category_count = len({
    expense.get("category", "Other")
    for expense in expenses
})

report_categories_value = ctk.CTkLabel(
    report_categories_card,
    text=str(category_count),
    font=("Arial", 26, "bold"),
    text_color="#FBBF24"
)

report_categories_value.pack(
    anchor="w",
    padx=20
)
# Reports - Donut Chart Container

donut_chart_frame = ctk.CTkFrame(
    content,
    fg_color=("#FFFFFF", "#1E293B"),
    corner_radius=15,
    height=350
)

donut_chart_frame.grid(
    row=14,
    column=0,
    columnspan=2,
    sticky="ew",
    padx=10,
    pady=15
)

donut_chart_frame.grid_propagate(False)
# Donut Chart Title

donut_chart_title = ctk.CTkLabel(
    donut_chart_frame,
    text="Expenses by Category",
    font=("Arial", 20, "bold"),
    text_color=("#0F172A", "#F8FAFC")
)

donut_chart_title.pack(
    anchor="w",
    padx=20,
    pady=(20, 5)
)
# Create Donut Chart

report_figure = Figure(
    figsize=(9, 3),
    dpi=100
)

report_ax = report_figure.add_subplot(111)

chart_bg = "#FFFFFF" if ctk.get_appearance_mode() == "Light" else "#1E293B"
report_figure.patch.set_facecolor(chart_bg)
report_ax.set_facecolor(chart_bg)
chart_text = "#0F172A" if ctk.get_appearance_mode() == "Light" else "#F8FAFC"

# Calculate expenses by category

expenses = load_expenses()
category_totals = {}

for expense in expenses:
    category = expense.get("category", "Other")
    amount = float(expense.get("amount", 0))

    category_totals[category] = (
        category_totals.get(category, 0) + amount
    )

# Draw Donut Chart

if category_totals:

    categories = list(category_totals.keys())
    amounts = list(category_totals.values())

    colors = [
        "#14B8A6",
        "#A78BFA",
        "#38BDF8",
        "#FBBF24",
        "#F87171",
        "#FB923C"
    ]

    wedges, texts, autotexts = report_ax.pie(
        amounts,
        labels=categories,
        autopct="%1.1f%%",
        startangle=90,
        colors=colors[:len(categories)],
        wedgeprops={
            "width": 0.4,
            "edgecolor": chart_bg
        },
        textprops={
            "color": chart_text,
            "fontsize": 10
        },
        pctdistance=0.78
    )

    report_ax.legend(
        wedges,
        categories,
        loc="center left",
        bbox_to_anchor=(1, 0.5),
        facecolor=chart_bg,
        edgecolor=chart_text,
        labelcolor=chart_text
    )

else:

    report_ax.text(
        0.5,
        0.5,
        "No expense data available",
        color="#475569" if ctk.get_appearance_mode() == "Light" else "#94A3B8",
        fontsize=14,
        ha="center",
        va="center"
    )

report_ax.axis("equal")

# Display Chart in CustomTkinter

donut_canvas = FigureCanvasTkAgg(
    report_figure,
    master=donut_chart_frame
)

donut_canvas.draw()

donut_canvas.get_tk_widget().pack(
    fill="both",
    expand=True,
    padx=15,
    pady=10
)
# Reports - Monthly Trend Chart Container

trend_chart_frame = ctk.CTkFrame(
    content,
    fg_color=("#FFFFFF", "#1E293B"),
    corner_radius=15,
    height=350
)

trend_chart_frame.grid(
    row=15,
    column=0,
    columnspan=2,
    sticky="ew",
    padx=10,
    pady=15
)

trend_chart_frame.grid_propagate(False)
# Monthly Trend Chart Title

trend_chart_title = ctk.CTkLabel(
    trend_chart_frame,
    text="Monthly Expenses Trend",
    font=("Arial", 20, "bold"),
    text_color=("#0F172A", "#F8FAFC")
)

trend_chart_title.pack(
    anchor="w",
    padx=20,
    pady=(20, 5)
)
# Create Monthly Trend Chart

trend_figure = Figure(
    figsize=(9, 3),
    dpi=100
)

trend_ax = trend_figure.add_subplot(111)

trend_chart_bg = "#FFFFFF" if ctk.get_appearance_mode() == "Light" else "#1E293B"
trend_figure.patch.set_facecolor(trend_chart_bg)
trend_ax.set_facecolor(trend_chart_bg)

# Calculate monthly expenses

expenses = load_expenses()
monthly_totals = {}

for expense in expenses:

    date_text = expense.get("date", "")

    try:
        expense_date = datetime.strptime(
            date_text,
            "%Y-%m-%d"
        )

        month = expense_date.strftime("%Y-%m")

        amount = float(expense.get("amount", 0))

        monthly_totals[month] = (
            monthly_totals.get(month, 0) + amount
        )

    except (ValueError, TypeError):
        continue

# Draw Line Chart

if monthly_totals:

    months = sorted(monthly_totals.keys())

    amounts = [
        monthly_totals[month]
        for month in months
    ]

    trend_ax.plot(
        months,
        amounts,
        color="#14B8A6",
        marker="o",
        linewidth=2.5,
        markersize=8
    )

    trend_ax.set_ylabel(
        "Total Expenses ($)",
        color="#0F172A" if ctk.get_appearance_mode() == "Light" else "#F8FAFC"
    )

    trend_ax.tick_params(
        axis="both",
        colors="#0F172A" if ctk.get_appearance_mode() == "Light" else "#F8FAFC"
    )

    trend_ax.grid(
        True,
        color="#CBD5E1" if ctk.get_appearance_mode() == "Light" else "#334155",
        linestyle="--",
        alpha=0.6
    )

    trend_ax.tick_params(
        axis="x",
        rotation=30
    )

else:

    trend_ax.text(
        0.5,
        0.5,
        "No expense data available",
        color="#475569" if ctk.get_appearance_mode() == "Light" else "#94A3B8",
        fontsize=14,
        ha="center",
        va="center"
    )

trend_figure.tight_layout()

# Display Chart in CustomTkinter

trend_canvas = FigureCanvasTkAgg(
    trend_figure,
    master=trend_chart_frame
)

trend_canvas.draw()

trend_canvas.get_tk_widget().pack(
    fill="both",
    expand=True,
    padx=15,
    pady=10
)

expenses = load_expenses()

lowest_amount = min(
    [float(expense.get("amount", 0)) for expense in expenses],
    default=0
)

report_lowest_value = ctk.CTkLabel(
    report_lowest_card,
    text=f"${lowest_amount:,.2f}",
    font=("Arial", 26, "bold"),
    text_color="#38BDF8"
)

report_lowest_value.pack(
    anchor="w",
    padx=20
)
# Settings Section

settings_frame = ctk.CTkFrame(
    content,
    fg_color=("#FFFFFF", "#1E293B"),
    corner_radius=15
)

settings_frame.grid(
    row=16,
    column=0,
    columnspan=2,
    sticky="ew",
    padx=10,
    pady=15
)

settings_title = ctk.CTkLabel(
    settings_frame,
    text="Settings",
    font=("Arial", 22, "bold"),
    text_color=("#0F172A", "#F8FAFC")
)

settings_title.pack(
    anchor="w",
    padx=20,
    pady=20
)
appearance_label = ctk.CTkLabel(
    settings_frame,
    text="Appearance Mode",
    font=("Arial", 16, "bold"),
    text_color=("#0F172A", "#F8FAFC")
)

appearance_label.pack(
    anchor="w",
    padx=20,
    pady=(10, 5)
)
appearance_menu = ctk.CTkOptionMenu(
    settings_frame,
    values=["Dark", "Light", "System"],
    width=180,
    height=35,
    command=change_appearance
)

appearance_menu.set("Dark")

appearance_menu.pack(
    anchor="w",
    padx=20,
    pady=(5, 20)
)
# About Section

about_label = ctk.CTkLabel(
    settings_frame,
    text="About",
    font=("Arial", 16, "bold"),
    text_color=("#0F172A", "#F8FAFC")
)

about_label.pack(
    anchor="w",
    padx=20,
    pady=(15, 5)
)
about_description = ctk.CTkLabel(
    settings_frame,
    text="Expense Tracker\nVersion 1.0.0\nManage and track your expenses easily.",
    font=("Arial", 13),
    text_color=("#475569", "#94A3B8"),
    justify="left"
)

about_description.pack(
    anchor="w",
    padx=20,
    pady=(5, 20)
)

def refresh_reports():

    expenses = load_expenses()
    chart_bg = "#FFFFFF" if ctk.get_appearance_mode() == "Light" else "#1E293B"
    chart_text = "#0F172A" if ctk.get_appearance_mode() == "Light" else "#F8FAFC"

    # Update Report Cards

    amounts = [
        float(expense.get("amount", 0))
        for expense in expenses
    ]

    total = sum(amounts)
    highest = max(amounts, default=0)
    lowest = min(amounts, default=0)

    categories = {
        expense.get("category", "Other")
        for expense in expenses
    }

    report_total_value.configure(
        text=f"${total:,.2f}"
    )

    report_highest_value.configure(
        text=f"${highest:,.2f}"
    )

    report_lowest_value.configure(
        text=f"${lowest:,.2f}"
    )

    report_categories_value.configure(
        text=str(len(categories))
    )

    # Update Donut Chart

    report_ax.clear()
    report_figure.patch.set_facecolor(chart_bg)

    report_ax.set_facecolor(chart_bg)

    category_totals = {}

    for expense in expenses:

        category = expense.get("category", "Other")
        amount = float(expense.get("amount", 0))

        category_totals[category] = (
            category_totals.get(category, 0) + amount
        )

    if category_totals:

        chart_categories = list(category_totals.keys())
        chart_amounts = list(category_totals.values())

        colors = [
            "#14B8A6",
            "#A78BFA",
            "#38BDF8",
            "#FBBF24",
            "#F87171",
            "#FB923C"
        ]

        wedges, texts, autotexts = report_ax.pie(
            chart_amounts,
            labels=chart_categories,
            autopct="%1.1f%%",
            startangle=90,
            colors=colors[:len(chart_categories)],
            wedgeprops={
                "width": 0.4,
                "edgecolor": chart_bg
            },
           textprops={
    "color": chart_text,
    "fontsize": 10
},
        )

        report_ax.legend(
            wedges,
            chart_categories,
            loc="center right",
            bbox_to_anchor=(0.98, 0.5),
            facecolor=chart_bg,
            edgecolor=chart_text,
            labelcolor=chart_text
        )

    else:

        report_ax.text(
            0.5,
            0.5,
            "No expense data available",
            color=chart_text,
            fontsize=14,
            ha="center",
            va="center"
        )

    report_ax.axis("equal")
    donut_canvas.draw()

    # Update Monthly Trend Chart

    trend_ax.clear()
    trend_figure.patch.set_facecolor(chart_bg)
    trend_ax.set_facecolor(chart_bg)

    monthly_totals = {}

    for expense in expenses:

        try:
            expense_date = datetime.strptime(
                expense.get("date", ""),
                "%Y-%m-%d"
            )

            month = expense_date.strftime("%Y-%m")
            amount = float(expense.get("amount", 0))

            monthly_totals[month] = (
                monthly_totals.get(month, 0) + amount
            )

        except (ValueError, TypeError):
            continue

    if monthly_totals:

        months = sorted(monthly_totals.keys())

        monthly_amounts = [
            monthly_totals[month]
            for month in months
        ]

        trend_ax.plot(
            months,
            monthly_amounts,
            color="#14B8A6",
            marker="o",
            linewidth=2.5,
            markersize=8
        )

        trend_ax.set_ylabel(
            "Total Expenses ($)",
            color=chart_text
        )

        trend_ax.tick_params(
            axis="both",
             colors=chart_text
            )

        trend_ax.grid(
            True,
            color=chart_text,
            linestyle="--",
            alpha=0.3
        )

        trend_ax.tick_params(
            axis="x",
            rotation=30
        )

    else:

        trend_ax.text(
            0.5,
            0.5,
            "No expense data available",
            color=chart_text,
            fontsize=14,
            ha="center",
            va="center"
        )

    trend_figure.tight_layout()
    trend_canvas.draw()
def refresh_categories():

    expenses = load_expenses()

    category_data = {}

    for expense in expenses:

        category = expense.get("category", "Other")
        amount = float(expense.get("amount", 0))

        if category not in category_data:
            category_data[category] = {
                "total": 0,
                "count": 0
            }

        category_data[category]["total"] += amount
        category_data[category]["count"] += 1

    # Remove old category rows

    for widget in categories_frame.winfo_children():
        if int(widget.grid_info()["row"]) > 0:
            widget.destroy()

    # Display category data

    for row_index, (category, data) in enumerate(
        category_data.items(),
        start=1
    ):

        values = [
            category,
            f"${data['total']:,.2f}",
            str(data["count"])
        ]

        for column_index, value in enumerate(values):

            label = ctk.CTkLabel(
                categories_frame,
                text=value,
                font=("Arial", 13),
                text_color=("#1E293B", "#F8FAFC")
            )

            label.grid(
                row=row_index,
                column=column_index,
                padx=15,
                pady=12,
                sticky="w"
            )
# Initialize dashboard
refresh_dashboard()

# Start application
app.mainloop()

