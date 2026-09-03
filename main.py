from datetime import datetime
from dateutil.relativedelta import relativedelta
import json
import os

from rich.console import Console
from rich.panel import Panel
from rich.table import Table
from rich.prompt import Prompt


FILE_PATH = "data/members.json"

console = Console()


# -----------------------------
# FILE HANDLING
# -----------------------------

def load_members():
    with open(FILE_PATH, "r") as file:
        return json.load(file)


def save_members(members):
    with open(FILE_PATH, "w") as file:
        json.dump(members, file, indent=4)


# -----------------------------
# CREATE
# -----------------------------

def add_member():
    
    console.print(
        Panel(
            "[bold cyan]➕ ADD NEW MEMBER[/bold cyan]",
            border_style="cyan"
        )
    )

    member_id = Prompt.ask("Member ID")
    name = Prompt.ask("Name")
    age = Prompt.ask("Age")
    phone = Prompt.ask("Phone")

    # -----------------------------
    # MEMBERSHIP PLAN
    # -----------------------------

    console.print("\n[bold]Select Membership Plan:[/bold]")
    console.print("1. Monthly")
    console.print("2. Quarterly")
    console.print("3. Half-Yearly")
    console.print("4. Yearly")

    plan_choice = Prompt.ask(
        "Enter choice",
        choices=["1", "2", "3", "4"]
    )

    if plan_choice == "1":
        plan = "Monthly"
        months = 1

    elif plan_choice == "2":
        plan = "Quarterly"
        months = 3

    elif plan_choice == "3":
        plan = "Half-Yearly"
        months = 6

    else:
        plan = "Yearly"
        months = 12

    # -----------------------------
    # START DATE
    # -----------------------------

    while True:

        start_date = Prompt.ask(
            "Start Date (YYYY-MM-DD)"
        )

        try:
            start = datetime.strptime(
                start_date,
                "%Y-%m-%d"
            )
            break

        except ValueError:
            console.print(
                "[bold red]Invalid date! "
                "Please use YYYY-MM-DD.[/bold red]"
            )

    # -----------------------------
    # EXPIRY DATE
    # -----------------------------

    expiry = start + relativedelta(
        months=months
    )

    expiry_date = expiry.strftime(
        "%Y-%m-%d"
    )

    # -----------------------------
    # PAYMENT STATUS
    # -----------------------------

    console.print("\n[bold]Select Payment Status:[/bold]")
    console.print("1. Paid")
    console.print("2. Pending")

    payment_choice = Prompt.ask(
        "Enter choice",
        choices=["1", "2"]
    )

    if payment_choice == "1":
        payment_status = "Paid"
    else:
        payment_status = "Pending"

    # -----------------------------
    # CREATE MEMBER
    # -----------------------------

    member = {
        "member_id": member_id,
        "name": name,
        "age": age,
        "phone": phone,
        "plan": plan,
        "start_date": start_date,
        "expiry_date": expiry_date,
        "payment_status": payment_status
    }

    # -----------------------------
    # SAVE MEMBER
    # -----------------------------

    members = load_members()

    members.append(member)

    save_members(members)

    console.print(
        Panel(
            f"[bold green]✓ Member Added Successfully![/bold green]\n\n"
            f"Member ID   : {member_id}\n"
            f"Name        : {name}\n"
            f"Plan        : {plan}\n"
            f"Start Date  : {start_date}\n"
            f"Expiry Date : {expiry_date}\n"
            f"Payment     : {payment_status}",
            title="GymTrack",
            border_style="green"
        )
    )
# -----------------------------
# READ
# -----------------------------

def view_members():
    members = load_members()

    if not members:
        console.print(
            Panel(
                "No members found.",
                title="Members",
                border_style="yellow"
            )
        )
        return

    table = Table(
        title="🏋️ Gym Members",
        show_lines=True
    )

    table.add_column("Member ID", style="cyan")
    table.add_column("Name")
    table.add_column("Age")
    table.add_column("Phone")

    for member in members:
        table.add_row(
            member["member_id"],
            member["name"],
            member["age"],
            member["phone"]
        )

    console.print(table)


# -----------------------------
# UPDATE
# -----------------------------

def update_member():
    member_id = Prompt.ask("Enter Member ID to update")

    members = load_members()

    for member in members:

        if member["member_id"] == member_id:

            console.print("\n[bold yellow]Current Details[/bold yellow]")
            console.print(f"Name  : {member['name']}")
            console.print(f"Age   : {member['age']}")
            console.print(f"Phone : {member['phone']}")

            member["name"] = Prompt.ask(
                "Enter new name",
                default=member["name"]
            )

            member["age"] = Prompt.ask(
                "Enter new age",
                default=member["age"]
            )

            member["phone"] = Prompt.ask(
                "Enter new phone",
                default=member["phone"]
            )

            save_members(members)

            console.print(
                "\n[bold green]✓ Member updated successfully![/bold green]"
            )

            return

    console.print("\n[bold red]✗ Member not found.[/bold red]")


# -----------------------------
# DELETE
# -----------------------------

def delete_member():
    member_id = Prompt.ask("Enter Member ID to delete")

    members = load_members()

    for member in members:

        if member["member_id"] == member_id:

            members.remove(member)
            save_members(members)

            console.print(
                "\n[bold green]✓ Member deleted successfully![/bold green]"
            )

            return

    console.print("\n[bold red]✗ Member not found.[/bold red]")
# ==============================
# EXPORT TO EXCEL
# ==============================

def export_to_excel():

    from openpyxl import Workbook

    members = load_members()

    if not members:
        console.print(
            Panel(
                "No members available to export.",
                title="Excel Export",
                border_style="yellow"
            )
        )
        return

    # Create exports folder
    if not os.path.exists("exports"):
        os.makedirs("exports")

    file_path = "exports/gym_members.xlsx"

    # Create Excel workbook
    workbook = Workbook()
    sheet = workbook.active
    sheet.title = "Gym Members"

    # Excel headings
    headers = [
        "Member ID",
        "Name",
        "Age",
        "Phone",
        "Plan",
        "Start Date",
        "Expiry Date",
        "Payment Status"
    ]

    sheet.append(headers)

    # Add member data
    for member in members:
        sheet.append([
            member["member_id"],
            member["name"],
            member["age"],
            member["phone"],
            member["plan"],
            member["start_date"],
            member["expiry_date"],
            member["payment_status"]
        ])

    # Adjust column widths
    for column in sheet.columns:

        max_length = 0
        column_letter = column[0].column_letter

        for cell in column:
            if cell.value:
                max_length = max(
                    max_length,
                    len(str(cell.value))
                )

        sheet.column_dimensions[column_letter].width = (
            max_length + 3
        )

    # Save Excel file
    workbook.save(file_path)

    console.print(
        Panel(
            f"[bold green]✓ Excel file created successfully![/bold green]\n\n"
            f"Location: {file_path}",
            title="📊 Excel Export",
            border_style="green"
        )
    )

# -----------------------------
# CREATE DATA FILE
# -----------------------------

if not os.path.exists("data"):
    os.makedirs("data")

if not os.path.exists(FILE_PATH):
    with open(FILE_PATH, "w") as file:
        json.dump([], file)


# -----------------------------
# MAIN MENU
# -----------------------------

while True:

    console.clear()

    console.print(
        Panel(
            "[bold cyan]🏋️ GYMTRACK[/bold cyan]\n"
            "Gym Member Management System",
            title="Welcome",
            border_style="cyan"
        )
    )

    console.print("\n[bold]Main Menu[/bold]")
    console.print("[cyan]1.[/cyan] 👤 Add Member")
    console.print("[cyan]2.[/cyan] 📋 View Members")
    console.print("[cyan]3.[/cyan] ✏️  Update Member")
    console.print("[cyan]4.[/cyan] 🗑️  Delete Member")
    console.print("[cyan]5.[/cyan] 📊 Export to Excel")
    console.print("[cyan]6.[/cyan] 🚪 Exit")

    choice = Prompt.ask(
        "\nChoose an option",
        choices=["1", "2", "3", "4", "5","6"]
    )

    if choice == "1":
        add_member()

    elif choice == "2":
        view_members()

    elif choice == "3":
        update_member()

    elif choice == "4":
        delete_member()

    elif choice == "5":
        export_to_excel()

    elif choice == "6":
        console.print(
        "\n[bold cyan]💪 Thanks for using GymTrack![/bold cyan]"
    )

        break

    Prompt.ask("\nPress Enter to continue", default="")