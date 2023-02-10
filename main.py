import secrets
import string
import subprocess
import win32com.client as win32
import random


def main():
    user_dict = {'username1': 'email assosiated', 'username2': 'another email'}
    for user in user_dict:
        new_password = create_password()
        if change_password(user, new_password) != -1:
            email_new_password(new_password, user_dict[user], user)


def create_password():
    special_characters = '!@#$%^&*+'
    all_variables = string.ascii_letters + string.digits + special_characters
    new_password = random.choice(string.ascii_lowercase) + random.choice(string.ascii_uppercase) + random.choice(
        string.digits) + random.choice(special_characters)
    for i in range(6):
        new_letter = secrets.choice(all_variables)
        new_password += ''.join(new_letter)
        all_variables.replace(new_letter, '')
    list_of_letters = list(new_password)
    random.SystemRandom().shuffle(list_of_letters)
    new_password = ''.join(list_of_letters)
    return new_password


def change_password(username, newpass):
    new_password_string = f'Set-ADAccountPassword -Identity {username} -Reset -NewPassword (ConvertTo-SecureString ' \
                          f'-AsPlainText "{newpass}" -Force)'
    pass_created = powershell_run(new_password_string)
    if pass_created.returncode != 0:
        email_tech(f"An error occurred creating a password {pass_created.stderr}", username)
        return -1
    else:
        return newpass


def email_tech(error, username):
    email = win32.Dispatch('outlook.application').CreateItem(0)
    Body = f'<hl>  <font size="+2">{error} for user {username} </font></hl>'
    email.to = 'tech@supportemail.com'
    email.Subject = 'Password Creation Script'
    email.htmlBody = Body
    email.Display()
    email.Send()


def email_new_password(new_password, user_email, username):
    email = win32.Dispatch('outlook.application').CreateItem(0)
    Body = f'<hl>  <font size="+2">The user {username}, has been set to this new password. {new_password} <br> If you ' \
           f'are having any problems logging in please submit a ticket <br> Thanks </font></hl>'
    email.to = user_email
    email.Subject = 'New Password for you user account'
    email.htmlBody = Body
    email.Display()
    email.Send()


def powershell_run(cmd):
    output = subprocess.run(["powershell", cmd], capture_output=True)
    return output


if __name__ == '__main__':
    main()
