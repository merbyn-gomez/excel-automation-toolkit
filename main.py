import xlwings as xw
import os


def copy_sheets(source_path, destination_workbook):
    """
    Copy worksheets from a source workbook into a destination workbook.
    """

    source_workbook = xw.Book(source_path)

    # Generate sheet name based on file name
    base_name = os.path.basename(source_path).split("_")[0]

    for sheet in source_workbook.sheets:
        copied_sheet = sheet.copy(after=destination_workbook.sheets[-1])

        # Rename copied sheet
        copied_sheet.name = base_name

    source_workbook.close()


def main():

    # Destination workbook path
    destination_path = input(
        "Enter destination workbook path (.xlsm): "
    )

    # Open or create destination workbook
    if os.path.exists(destination_path):
        destination_workbook = xw.Book(destination_path)
    else:
        destination_workbook = xw.Book()
        destination_workbook.save(destination_path)

    # Source files
    source_files = input(
        "Enter source Excel files separated by commas: "
    ).split(",")

    # Remove spaces
    source_files = [file.strip() for file in source_files]

    # Copy sheets
    for file in source_files:
        copy_sheets(file, destination_workbook)

    # Save and close
    destination_workbook.save()
    destination_workbook.close()

    print("Sheets copied successfully.")


if __name__ == "__main__":
    main()
