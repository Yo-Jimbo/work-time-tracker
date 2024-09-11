# Work Time Tracker

## Description
This is an application designed to help users log hours worked on different clients in a time tracking software. It helps automating the process of selecting clients, entering notes, and confirming the hours worked. The website hosting the time tracking software has been redacted as "https://gestionale.com/", as such the code is provided solely for demonstration purposes.

## Requirements
- **Python 3.x**
- **Libraries**: requests, bs4, openpyxl, tkinter, plyer, selenium
- **If using Windows**: putting the "logo-gestionale.ico" file in the same folder as the .py file is required if on Windows operating system.

## Usage
- **Login**: Enter your credentials (username and password) in the login window.
- **Notifications**: After successfully logging in, the script will send a notification at minutes 00 and 30 of every hour, informing that it is now possible to register and associate the next half-hour to a client.
- **Client selection**: Select a client and enter any notes in the main window.
- **Confirm the half-hour slot**: Click on the "Conferma" ("Confirm") button to register and store the hours worked for the selected client in the "AutoGestionale.xlsx" file.
- **Check for previous missing time slots**: If there are missing slots (after logging in later than at 9AM or after not registering a previous half-hour slot, click on "Compila gli slot mancanti" ("Fill the missing slots")to fill them.
- **Registration on the TIme Tracker software**: Once all hours are registered, click on "Apri Gestionale" ("Open Time Tracker") to open the Time Tracker's website and register the hours there.

## Features
- **Updated and customizable**: Selection of clients updated according to the list of clients available to the user and entry of notes.
- **No time lost**: Handling of missing half-hour slots due to later log-in or after not registering previous time slots.
- **User-friendly**: Simple and intuitive interface with notifications and alerts.

