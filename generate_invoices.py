import openpyxl
from datetime import datetime
from pathlib import Path


current_datetime = datetime.now()


def generate_invoce_pdf(
    client_email: str,
    client_name: str,
    client_address: str,
    sale_description: str,
    client_cost: str,
    invoice_number: str,
    date_of_sale: str,
    due_date: str,
):
    cell_data = {
        "B10": client_name,
        "D10": client_address,
        "B16": sale_description,
        "E16": client_cost,
        "F9": invoice_number,
        "F10": date_of_sale,
        "F11": due_date,
        "F17": client_cost,
    }
    template_workbook = openpyxl.load_workbook("Munder Difflin.xlsx")
    sheet = template_workbook["Invoice Template"]

    for key in cell_data:
        sheet[key] = cell_data[key]

    template_workbook.save(f"invoices/{invoice_number}.xlsx")

    # TODO: functionality to send email
    print(f"Sending Invoice to: {client_email}")


def generate_invoice_number() -> str:
    # sample invoice numebr: DM-10152020-0001
    sequence_id = "0000"
    sequence_tracker_filename = "sequence_tracker.txt"
    latest_invoice_number = f"DM-{current_datetime.strftime(r'%m%d%Y')}-{sequence_id}"
    sequence_tracker_file = Path(sequence_tracker_filename)
    if sequence_tracker_file.is_file():
        with open(sequence_tracker_filename) as f:
            if invoice_numbers := [line.rstrip() for line in f]:
                latest_invoice_number = sorted(
                    invoice_numbers,
                    key=lambda x: (
                        datetime.strptime(x.split("-")[1], "%m%d%Y"),
                        x.split("-")[-1],
                    ),
                    reverse=True,
                )[0]

    if latest_invoice_number.split("-")[1] == current_datetime.strftime(r"%m%d%Y"):
        latest_sequence_id = latest_invoice_number.split("-")[-1]
        sequence_id = format(int(latest_sequence_id) + 1, "04d")
        latest_invoice_number = (
            f"DM-{current_datetime.strftime(r'%m%d%Y')}-{sequence_id}"
        )

    with open(sequence_tracker_file, "a") as f:
        f.write(f"{latest_invoice_number}\n")

    return latest_invoice_number


if __name__ == "__main__":
    # The client's email address
    client_email = input("Client's Email: ")
    # Client's name
    client_name = input("Client's Name: ")
    # Clients address
    client_address = input("Client's Address: ")
    # The description of the sale
    sale_description = input("Description of sale: ") or ""
    # The cost to the client
    client_cost = input("Client's Cost: $")
    # An invoice number
    invoice_number = generate_invoice_number()
    # Today's date
    date_of_sale = current_datetime.strftime(r"%m/%d/%Y")
    # A due date
    due_date = input("Due Date (mm/dd/yyyy): ")
    user_input = {
        key: eval(key)
        for key in [
            "client_email",
            "client_name",
            "client_address",
            "sale_description",
            "client_cost",
            "invoice_number",
            "date_of_sale",
            "due_date",
        ]
    }
    generate_invoce_pdf(**user_input)