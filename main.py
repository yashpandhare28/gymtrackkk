from datetime import datetime
import json
import os

from rich.console import Console
from rich.panel import Panel
from rich.table import Table
from rich.prompt import Prompt
from openpyxl import Workbook


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

# -----------------------------
# CREATE
# -----------------------------

def add_member():

    console.print(
        Panel(
            "[bold cyan]ADD NEW MEMBER[/bold cyan]",
            border_style="cyan"
        )
    )

    member_id = Prompt.ask("Enter Member ID")
    name = Prompt.ask("Name")
    age = Prompt.ask("Age")
    phone = Prompt.ask("Phone")

    # Membership Plan
    console.print("\n[bold cyan]Select Membership Plan[/bold cyan]")
    console.print("1. 1 Month")
    console.print("2. 6 Months")
    console.print("3. 1 Year")

    plan_choice = Prompt.ask(
        "Choose plan",
        choices=["1", "2", "3"]
    )

    if plan_choice == "1":
        plan = "1 Month"
    elif plan_choice == "2":
        plan = "6 Months"
    else:
        plan = "1 Year"

    # Start Date
    start_date = Prompt.ask("Start Date (YYYY-MM-DD)")

    # Payment Status
    console.print("\n[bold cyan]Select Payment Status[/bold cyan]")
    console.print("1. Paid")
    console.print("2. Pending")

    payment_choice = Prompt.ask(
        "Choose payment status",
        choices=["1", "2"]
    )

    if payment_choice == "1":
        payment = "Paid"
    else:
        payment = "Pending"

    # Create member
    member = {
        "member_id": member_id,
        "name": name,
        "age": age,
        "phone": phone,
        "plan": plan,
        "start_date": start_date,
        "payment_status": payment
    }

    members = load_members()
    members.append(member)
    save_members(members)

    console.print(
        Panel(
            "[bold green]Member Added Successfully![/bold green]",
            border_style="green"
        )
    )

# -----------------------------
# READ
# -----------------------------

def view_members():
    members = load_members()

    if not members:
        console.print("[yellow]No members found.[/yellow]")
        return

    table = Table(title="Gym Members")

    table.add_column("ID", style="cyan")
    table.add_column("Name")
    table.add_column("Age")
    table.add_column("Phone")
    table.add_column("Plan")
    table.add_column("Start Date")
    table.add_column("Payment")

    for member in members:
        table.add_row(
            member["member_id"],
            member["name"],
            member["age"],
            member["phone"],
            member["plan"],
            member["start_date"],
            member["payment_status"]
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

            member["name"] = Prompt.ask(
                "New Name",
                default=member["name"]
            )

            member["age"] = Prompt.ask(
                "New Age",
                default=member["age"]
            )

            member["phone"] = Prompt.ask(
                "New Phone",
                default=member["phone"]
            )

            member["plan"] = Prompt.ask(
                "New Plan",
                default=member["plan"]
            )

            member["payment_status"] = Prompt.ask(
                "Payment Status",
                choices=["Paid", "Pending"],
                default=member["payment_status"]
            )

            save_members(members)

            console.print(
                "[green]Member updated successfully![/green]"
            )
            return

    console.print("[red]Member not found.[/red]")


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
                "[green]Member deleted successfully![/green]"
            )
            return

    console.print("[red]Member not found.[/red]")


# -----------------------------
# EXPORT TO EXCEL
# -----------------------------

def export_to_excel():
    members = load_members()

    if not members:
        console.print("[yellow]No members to export.[/yellow]")
        return

    os.makedirs("exports", exist_ok=True)

    workbook = Workbook()
    sheet = workbook.active
    sheet.title = "Gym Members"

    headers = [
        "Member ID",
        "Name",
        "Age",
        "Phone",
        "Plan",
        "Start Date",
        "Payment Status"
    ]

    sheet.append(headers)

    for member in members:
        sheet.append([
            member["member_id"],
            member["name"],
            member["age"],
            member["phone"],
            member["plan"],
            member["start_date"],
            member["payment_status"]
        ])

    workbook.save("exports/gym_members.xlsx")

    console.print(
        Panel(
            "[bold green]Excel file created successfully![/bold green]\n"
            "Location: exports/gym_members.xlsx",
            title="Excel Export",
            border_style="green"
        )
    )


# -----------------------------
# MAIN MENU
# -----------------------------

def main():
    os.makedirs("data", exist_ok=True)

    if not os.path.exists(FILE_PATH):
        save_members([])

    while True:
        console.clear()

        console.print(
            Panel(
                "[bold cyan]GYMTRACK[/bold cyan]\n"
                "Gym Member Management System",
                title="Welcome",
                border_style="cyan"
            )
        )

        console.print("\n[bold]Main Menu[/bold]")
        console.print("1. Add Member")
        console.print("2. View Members")
        console.print("3. Update Member")
        console.print("4. Delete Member")
        console.print("5. Export to Excel")
        console.print("6. Exit")

        choice = Prompt.ask(
            "\nChoose an option",
            choices=["1", "2", "3", "4", "5", "6"]
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
                "\n[bold cyan]Thanks for using GymTrack![/bold cyan]"
            )
            break

        Prompt.ask("\nPress Enter to continue", default="")


if __name__ == "__main__":
    main()