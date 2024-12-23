▎Certificates and Email Sending

This project is designed to create personalized certificates in PDF format and send them via email. It utilizes the pandas, aspose.pdf, and smtplib libraries to perform its tasks.

▎Description

The script performs the ▎Certificates and Email Sending

This project is designed to create personalized certificates in PDF format and send them via email. It utilizes the pandas, aspose.pdf, and smtplib libraries to perform its tasks.

▎Description

The script performs the following functions:

1. Reading Data: Reads participant names and their email addresses from an Excel file.

2. Creating Certificates: Generates PDF certificates based on a template, replacing the text "Full Name" with the participant's name.

3. Sending Emails: Sends the created certificates to the specified email addresses.

▎Installation

To run the script, you need to install the following libraries:

pip install pandas aspose-pdf


▎Usage

1. Prepare a file named data1.xlsx with two columns:

   • The first column: full names of participants.

   • The second column: email addresses.

2. Create a text file named text.txt containing the text that will be included in the certificate.

3. Replace file paths:

   • The path to the certificate template in the template_path variable.

   • The path to the Excel file in the excel_file variable.

4. Ensure that you replace the credentials and SMTP server settings in the send_email function.

5. Run the script:

python your_script.py


▎Notes

• Make sure you have access to an SMTP server and the correct credentials for sending emails.

• The certificate template should contain the text "Full Name," which will be replaced with the participant's name.

• The script automatically creates certificates and sends them via email.

▎License

This project is licensed under the MIT License.

▎Contribution

If you would like to make changes or improvements, please create a fork of the repository and submit a pull request.

---

If you have any questions or issues, feel free to reach out for assistance!following functions:

1. Reading Data: Reads participant names and their email addresses from an Excel file.

2. Creating Certificates: Generates PDF certificates based on a template, replacing the text "Full Name" with the participant's name.

3. Sending Emails: Sends the created certificates to the specified email addresses.

▎Installation

To run the script, you need to install the following libraries:

pip install pandas aspose-pdf


▎Usage

1. Prepare a file named data1.xlsx with two columns:

   • The first column: full names of participants.

   • The second column: email addresses.

2. Create a text file named text.txt containing the text that will be included in the certificate.

3. Replace file paths:

   • The path to the certificate template in the template_path variable.

   • The path to the Excel file in the excel_file variable.

4. Ensure that you replace the credentials and SMTP server settings in the send_email function.

5. Run the script:

python your_script.py


▎Notes

• Make sure you have access to an SMTP server and the correct credentials for sending emails.

• The certificate template should contain the text "Full Name," which will be replaced with the participant's name.

• The script automatically creates certificates and sends them via email.

▎License

This project is licensed under the MIT License.

▎Contribution

If you would like to make changes or improvements, please create a fork of the repository and submit a pull request.

---

If you have any questions or issues, feel free to reach out for assistance!