import random
import string
import openpyxl

# Program to generate secure passwords, store them along with website addresses
# and usernames in an Excel file, and allow the user to continue
# generating passwords until they choose to exit.
def generate_password(length=12):
    # Define the characters to use for password generation
    characters = string.ascii_letters + string.digits + string.punctuation
    password = ''.join(random.choice(characters) for _ in range(length))
    return password
def save_password_to_excel(websiteaddress, username, password):
    filename = "C:\\Users\\User\\Desktop\\Web\\generated_passwords.xlsx"

    try:
        # Try loading the existing Excel workbook
        workbook = openpyxl.load_workbook(filename)
    except FileNotFoundError:
        # If the file doesn't exist, create a new workbook and add a header row
        workbook = openpyxl.Workbook()
        sheet = workbook.active
        sheet.append(["Website", "Username", "Password"])

    sheet = workbook.active
    sheet.append([websiteaddress, username, password])

    workbook.save(filename)
    print(f"Password saved to {filename}")

if __name__ == "__main__":
    while True:
        # Main loop to generate passwords
        password_length_input = input("Enter the desired password length (enter 'exit' to quit): ")
        if password_length_input.lower() == 'exit':
            break

        try:
            password_length = int(password_length_input)
        except ValueError:
            print("Invalid input. Please enter a valid password length.")
            continue

        website = input("Enter the website for which the password is generated: ")
        if '.' not in website:
            print("Please enter a proper web address.")
            continue

        username = input("Enter the username associated with the password: ")

        generated_password = generate_password(password_length)
        print("Generated Password:", generated_password)

        save_password_to_excel(website, username, generated_password)
