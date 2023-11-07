
import PyPDF2
import sys
import time

print(sys.argv[0])
if len(sys.argv) < 4:
     print('You must pass input_file, output_file and password', file=sys.stderr)
     quit()
#file name is path of pdf file
input_filename = sys.argv[1] 
output_filename = sys.argv[2] 
password = sys.argv[3]


# Open the PDF file in read mode
with open(input_filename, "rb") as f:
    # Create a PDF reader object
    reader = PyPDF2.PdfReader(f)

    # Create a PDF writer object
    writer = PyPDF2.PdfWriter()

    # Add each page of the PDF to the writer object
    for i in range(len(reader.pages)):
        writer.add_page(reader.pages[i])

    # Encrypt the PDF with the given password and permissions
    writer.encrypt(user_password="", owner_password=password, use_128bit=True, permissions_flag=0b0100)
    #writer.encrypt(password, use_128bit=True, perm=permissions)

    # Save the encrypted PDF to a new file
    with open(output_filename, "wb") as f:
        writer.write(f)
        
print(output_filename + " file created successfully")

