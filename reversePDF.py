import PyPDF2
import win32api
import win32print
import os


# Function to reverse PDF pages
def reverse_pdf(input_pdf, output_pdf):
    with open(input_pdf, 'rb') as file:
        reader = PyPDF2.PdfReader(file)
        writer = PyPDF2.PdfWriter()

        # Reverse the order of pages
        num_pages = len(reader.pages)
        for page_num in range(161, 80, -2):
            page = reader.pages[page_num]
            writer.add_page(page)

        # Write the reversed pages to a new PDF
        with open(output_pdf, 'wb') as output_file:
            writer.write(output_file)


# Reverse the pages of the input PDF
input_pdf = 'input.pdf'
output_pdf = 'reversed_output.pdf'
reverse_pdf(input_pdf, output_pdf)

# Print the reversed PDF
# printer_name = win32print.GetDefaultPrinter()
# win32api.ShellExecute(
#     0,
#     "print",
#     output_pdf,
#     f'/d:"{printer_name}"',
#     ".",
#     0
# )

# Clean up the reversed PDF file if desired
# os.remove(output_pdf)
